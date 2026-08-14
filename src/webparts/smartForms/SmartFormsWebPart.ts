import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/views';

// Type-only: IReadonlyTheme is an interface, so this import is erased at
// compile time and @microsoft/sp-component-base never becomes a runtime AMD
// dependency of this bundle. The theme itself arrives via onThemeChanged,
// which BaseClientSideWebPart calls for us — consuming ThemeProvider from
// the service scope is not required and adds an external for no benefit.
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import { SharePointService } from './services/SharePointService';
import { SmartForms } from './components/SmartForms';
import { ErrorBoundary } from './components/ErrorBoundary';
import { debugLog, logError } from './utils/debug';

export interface ISmartFormsWebPartProps {
  listId: string;
  formDefinition: string;
  newListName: string;
}

export default class SmartFormsWebPart extends BaseClientSideWebPart<ISmartFormsWebPartProps> {
  private spService: SharePointService;
  private availableLists: IPropertyPaneDropdownOption[] = [];
  private listsLoaded = false;
  private paneStartDeferred = false;
  private initialized = false;
  private createListStatus = '';
  private currentTheme: IReadonlyTheme | undefined;

  public async onInit(): Promise<void> {
    await super.onInit();
    const sp = spfi().using(SPFx(this.context));
    this.spService = new SharePointService(sp);
    // From here on this.properties / this.displayMode are safe to read, so
    // onThemeChanged is allowed to trigger repaints.
    this.initialized = true;
    debugLog('onInit complete', {
      listId: this.properties.listId,
      displayMode: DisplayMode[this.displayMode]
    });
    // The pane may have tried to open before we were ready (see
    // onPropertyPaneConfigurationStart); pick up the work it had to skip.
    if (this.paneStartDeferred) {
      this.paneStartDeferred = false;
      this.loadAvailableLists();
    }
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    this.currentTheme = theme;
    // SPFx raises the first theme change *before* onInit, while this.properties
    // and this.displayMode are still unavailable — touching either from render()
    // at that point throws inside the host's theme bootstrap, which aborts the
    // whole component load and surfaces only as "ERROR: [object Object]".
    // There is nothing to repaint yet anyway: SPFx renders once onInit resolves,
    // and currentTheme above is already stored for that first paint.
    if (!this.initialized) {
      return;
    }
    this.render();
  }

  public render(): void {
    // This guard must come first and must not touch this.properties or
    // this.displayMode: render() can be reached before SPFx has initialized
    // the component, and those getters throw at that point. (Note that a
    // debugLog(...) call here is not safe either — its argument object is
    // evaluated eagerly even when debug logging is switched off.)
    if (!this.spService) {
      return;
    }
    debugLog('render', { listId: this.properties.listId, displayMode: DisplayMode[this.displayMode] });
    const element = React.createElement(
      ErrorBoundary,
      null,
      React.createElement(SmartForms, {
        listId: this.properties.listId,
        formDefinitionJson: this.properties.formDefinition,
        instanceId: this.context.instanceId,
        spService: this.spService,
        isEditMode: this.displayMode === DisplayMode.Edit,
        theme: this.currentTheme,
        onConfigure: () => this.context.propertyPane.open(),
        onFormDefinitionChange: (definitionJson: string) => {
          // kept as a lightweight mirror on the page property (useful if the
          // page is exported/copied); SharePointService.saveFormDefinition is
          // what actually makes the edit durable, independent of page save.
          this.properties.formDefinition = definitionJson;
        }
      })
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    // Adding the web part to a page auto-opens the property pane, and that can
    // happen before the async onInit has resolved — so spService may not exist
    // yet. Dereferencing it here throws inside SPFx's property-pane bootstrap,
    // where the exception surfaces only as an opaque error tile. Defer instead;
    // onInit picks this up as soon as the service is available.
    if (this.listsLoaded) {
      return;
    }
    if (!this.spService) {
      this.paneStartDeferred = true;
      return;
    }
    this.loadAvailableLists();
  }

  private loadAvailableLists(): void {
    this.spService
      .getAvailableLists()
      .then((lists) => {
        this.availableLists = lists.map((l) => ({ key: l.id, text: l.title }));
        this.listsLoaded = true;
        debugLog('getAvailableLists loaded', { count: lists.length });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        logError('getAvailableLists', error);
        this.availableLists = [];
        this.listsLoaded = true;
        this.createListStatus = 'The lists on this site could not be loaded.';
        this.context.propertyPane.refresh();
      });
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'listId' && oldValue !== newValue) {
      this.render();
    }
  }

  private onCreateListClick = (): void => {
    const name = (this.properties.newListName || '').trim();
    if (!name) {
      this.createListStatus = 'Enter a name for the new list first.';
      this.context.propertyPane.refresh();
      return;
    }
    this.createListStatus = 'Creating "' + name + '"…';
    this.context.propertyPane.refresh();

    this.spService
      .createList(name)
      .then((list) => {
        this.properties.listId = list.id;
        this.properties.newListName = '';
        this.availableLists = this.availableLists
          .concat([{ key: list.id, text: list.title }])
          .sort((a, b) => String(a.text).localeCompare(String(b.text)));
        this.createListStatus = 'List "' + list.title + '" created and selected.';
        this.context.propertyPane.refresh();
        this.render();
      })
      .catch((e) => {
        if (!this.spService.isDuplicateListNameError(e)) {
          logError('createList', e);
        }
        this.createListStatus = this.spService.isDuplicateListNameError(e)
          ? 'A list named "' + name + '" already exists — select it from the dropdown above.'
          : 'Could not create the list. Check that you have permission to create lists on this site.';
        this.context.propertyPane.refresh();
      });
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const listSelected = !!this.properties.listId;
    const selectedList = this.availableLists.filter((l) => l.key === this.properties.listId)[0];

    return {
      pages: [
        {
          header: {
            description:
              'Choose the SharePoint list that stores form responses, or create a new one. Everything else — questions, appearance, notifications — is edited directly on the page.'
          },
          groups: [
            {
              groupName: 'Response list',
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'Save responses to this list',
                  options: this.availableLists,
                  selectedKey: this.properties.listId,
                  disabled: !this.listsLoaded
                }),
                PropertyPaneLabel('listStatus', {
                  text: this.listsLoaded
                    ? listSelected && selectedList
                      ? 'Responses go to "' + selectedList.text + '".'
                      : 'Pick a list, or create one below.'
                    : 'Loading lists…'
                }),
                PropertyPaneTextField('newListName', {
                  label: 'Or create a new list',
                  placeholder: 'New list name'
                }),
                PropertyPaneButton('createList', {
                  text: 'Create list',
                  buttonType: PropertyPaneButtonType.Primary,
                  icon: 'Add',
                  onClick: this.onCreateListClick
                }),
                PropertyPaneLabel('createStatus', {
                  text: this.createListStatus
                })
              ]
            },
            {
              groupName: 'Tips',
              isCollapsed: true,
              groupFields: [
                PropertyPaneLabel('tipPublish', {
                  text:
                    'Questions and settings save automatically as you edit them — no need to save or publish the page.'
                }),
                PropertyPaneLabel('tipPermissions', {
                  text:
                    'Respondents need permission to add items to the response list. Check the list permissions if someone reports an error when submitting.'
                }),
                PropertyPaneLabel('tipShare', {
                  text:
                    'Use "Collect responses" on the page to create the list columns and get a shareable fill-in link.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
