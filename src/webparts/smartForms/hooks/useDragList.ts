import * as React from 'react';

/**
 * HTML5 drag-and-drop reordering for a flat list.
 *
 * Returns prop getters rather than rendering anything, so the same behavior
 * backs the ranking control, the designer canvas and the outline rail without
 * three copies of the dragover bookkeeping.
 *
 * `dropSide` tells the caller whether the pointer is in the top or bottom half
 * of the hovered row, which is what lets the indicator line appear above or
 * below rather than always snapping to one edge.
 */

export type DropSide = 'before' | 'after';

export interface IDragListState {
  dragIndex: number;
  overIndex: number;
  dropSide: DropSide;
}

export interface IDragHandlers {
  draggable: boolean;
  onDragStart: (event: React.DragEvent) => void;
  onDragOver: (event: React.DragEvent) => void;
  onDragLeave: (event: React.DragEvent) => void;
  onDrop: (event: React.DragEvent) => void;
  onDragEnd: (event: React.DragEvent) => void;
}

export interface IDragList {
  state: IDragListState;
  /** spread onto each row element */
  rowProps: (index: number) => IDragHandlers;
  /** true when this row is the one being dragged */
  isDragging: (index: number) => boolean;
  /** where the drop indicator belongs for this row, if anywhere */
  dropIndicator: (index: number) => DropSide | undefined;
  /** imperative reorder, for the keyboard up/down controls */
  move: (from: number, to: number) => void;
  cancel: () => void;
}

const IDLE: IDragListState = { dragIndex: -1, overIndex: -1, dropSide: 'before' };

export interface IUseDragListOptions {
  /** total rows currently rendered */
  count: number;
  /** called with the source and destination indices once a drop lands */
  onReorder: (from: number, to: number) => void;
  /** payload written to the drag event; defaults to the index */
  dataTransferText?: (index: number) => string;
  /** set false to disable dragging entirely (e.g. a locked question) */
  enabled?: boolean;
}

export const useDragList = (options: IUseDragListOptions): IDragList => {
  const { count, onReorder } = options;
  const enabled = options.enabled !== false;
  const [state, setState] = React.useState<IDragListState>(IDLE);

  // the handlers below read drag state during a native drag; a ref keeps them
  // reading the current value without rebuilding every row's props each move
  const stateRef = React.useRef<IDragListState>(state);
  stateRef.current = state;

  const cancel = React.useCallback((): void => setState(IDLE), []);

  const move = React.useCallback(
    (from: number, to: number): void => {
      if (from === to || from < 0 || to < 0 || from >= count || to >= count) {
        return;
      }
      onReorder(from, to);
    },
    [count, onReorder]
  );

  const rowProps = (index: number): IDragHandlers => ({
    draggable: enabled,
    onDragStart: (event: React.DragEvent) => {
      if (!enabled) {
        return;
      }
      setState({ dragIndex: index, overIndex: -1, dropSide: 'before' });
      event.dataTransfer.effectAllowed = 'move';
      // Firefox refuses to start a drag with no payload
      const text = options.dataTransferText ? options.dataTransferText(index) : String(index);
      try {
        event.dataTransfer.setData('text/plain', text);
      } catch {
        // some hosts lock the data store during dragstart; the drag still works
      }
    },
    onDragOver: (event: React.DragEvent) => {
      if (!enabled || stateRef.current.dragIndex < 0) {
        return;
      }
      event.preventDefault();
      event.dataTransfer.dropEffect = 'move';
      const bounds = (event.currentTarget as HTMLElement).getBoundingClientRect();
      const side: DropSide = event.clientY - bounds.top < bounds.height / 2 ? 'before' : 'after';
      const current = stateRef.current;
      if (current.overIndex !== index || current.dropSide !== side) {
        setState({ dragIndex: current.dragIndex, overIndex: index, dropSide: side });
      }
    },
    onDragLeave: () => {
      setState((prev) => (prev.overIndex === index ? { ...prev, overIndex: -1 } : prev));
    },
    onDrop: (event: React.DragEvent) => {
      if (!enabled) {
        return;
      }
      event.preventDefault();
      const { dragIndex, dropSide } = stateRef.current;
      setState(IDLE);
      if (dragIndex < 0) {
        return;
      }
      // dropping "after" row N means landing at N+1, but removing the source
      // first shifts everything below it up by one
      let to = dropSide === 'after' ? index + 1 : index;
      if (dragIndex < to) {
        to--;
      }
      move(dragIndex, to);
    },
    onDragEnd: () => setState(IDLE)
  });

  return {
    state,
    rowProps,
    isDragging: (index: number) => state.dragIndex === index,
    dropIndicator: (index: number) =>
      state.overIndex === index && state.dragIndex >= 0 && state.dragIndex !== index
        ? state.dropSide
        : undefined,
    move,
    cancel
  };
};

/** Move an item within a copy of `items`. */
export const reorder = <T,>(items: T[], from: number, to: number): T[] => {
  const next = items.slice();
  const moved = next.splice(from, 1)[0];
  next.splice(to, 0, moved);
  return next;
};
