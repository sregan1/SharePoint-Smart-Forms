import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { IReadonlyTheme, ThemeProvider } from '@microsoft/sp-component-base';
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
  private createListStatus = '';
  private currentTheme: IReadonlyTheme | undefined;

  protected async onInit(): Promise<void> {
    await super.onInit();
    debugLog('onInit starting', {
      version: this.manifest.version,
      listId: this.properties.listId,
      displayMode: DisplayMode[this.displayMode]
    });
    try {
      const sp = spfi().using(SPFx(this.context));
      this.spService = new SharePointService(sp);

      // The site theme drives every color in the UI. Without it the form renders
      // its own light palette on a dark-themed site and is unreadable. Read the
      // initial value here; the base class calls onThemeChanged for later updates.
      const themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
      this.currentTheme = themeProvider.tryGetTheme();
      debugLog('onInit complete', { hasTheme: !!this.currentTheme });
    } catch (error) {
      // onInit throwing is the classic "web part shows nothing, nothing in the
      // console" failure — SPFx's own handling of a rejected onInit doesn't
      // always surface a clear message, so log one explicitly here.
      logError('onInit', error);
      throw error;
    }
  }

  protected onThemeChanged(theme: IReadonlyTheme | undefined): void {
    this.currentTheme = theme;
    this.render();
  }

  public render(): void {
    debugLog('render', { listId: this.properties.listId, displayMode: DisplayMode[this.displayMode] });
    const element = React.createElement(
      ErrorBoundary,
      null,
      React.createElement(SmartForms, {
        listId: this.properties.listId,
        formDefinitionJson: this.properties.formDefinition,
        spService: this.spService,
        isEditMode: this.displayMode === DisplayMode.Edit,
        theme: this.currentTheme,
        onConfigure: () => this.context.propertyPane.open(),
        onFormDefinitionChange: (definitionJson: string) => {
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
    if (this.listsLoaded) {
      return;
    }
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
                    'The form is stored with the page, so save or publish the page after editing questions.'
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
