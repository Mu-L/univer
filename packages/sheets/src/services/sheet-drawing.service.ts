/**
 * Copyright 2023-present DreamNum Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import type { DrawingType, IDrawingSearch, IOtherTransform, ISize, Nullable } from '@univerjs/core';
import { Disposable } from '@univerjs/core';
import { createIdentifier } from '@wendellhu/redi';
import type { Observable } from 'rxjs';
import { Subject } from 'rxjs';

interface ICellPosition {
    column: number; // column number
    columnOffset: number; // column offset, unit is EMUs
    row: number; // row number
    rowOffset: number; // row offset, unit is EMUs
}


export interface ISheetDrawingPosition extends IOtherTransform {
    from: ICellPosition;
    to: ICellPosition;
}

export interface ISheetDrawingParam {
    sheetTransform: ISheetDrawingPosition;
    originSize: ISize;
}


export interface ISheetDrawingServiceParam extends IDrawingSearch, ISheetDrawingParam {

}

export interface ISheetDrawingServiceUpdateParam extends IDrawingSearch, Partial<ISheetDrawingParam> {

}


export interface ISheetDrawingMap {
    [unitId: string]: ISheetDrawingSubunitMap;
}

export interface ISheetDrawingSubunitMap {
    [subUnitId: string]: IISheetDrawingMapItem;
}

export interface IISheetDrawingMapItem {
    [drawingId: string]: ISheetDrawingServiceParam;
}

export interface ISheetDrawingService {
    addDrawing(param: ISheetDrawingServiceParam): void;
    batchAddDrawing(params: ISheetDrawingServiceParam[]): void;

    removeDrawing(param: IDrawingSearch): Nullable<ISheetDrawingServiceParam>;
    batchRemoveDrawing(params: IDrawingSearch[]): ISheetDrawingServiceParam[];

    focusDrawing(param: Nullable<IDrawingSearch>): void;
    getFocusDrawings(): ISheetDrawingServiceParam[];

    updateDrawing(param: ISheetDrawingServiceUpdateParam): void;
    batchUpdateDrawing(params: ISheetDrawingServiceUpdateParam[]): void;

    getDrawingMap(unitId: string, subUnitId: string, drawingType: DrawingType): Nullable<IISheetDrawingMapItem>;
    getDrawingItem(param: IDrawingSearch): Nullable<ISheetDrawingServiceParam>;

    add$: Observable<ISheetDrawingServiceParam[]>;
    remove$: Observable<ISheetDrawingServiceParam[]>;
    update$: Observable<ISheetDrawingServiceUpdateParam[]>;

}

export class SheetDrawingService extends Disposable implements ISheetDrawingService {
    private _drawingMap: { [drawingType: DrawingType]: ISheetDrawingMap } = {};

    private _focusDrawings: ISheetDrawingServiceParam[] = [];

    private _add$ = new Subject<ISheetDrawingServiceParam[]>();
    add$: Observable<ISheetDrawingServiceParam[]> = this._add$.asObservable();

    private _remove$ = new Subject<ISheetDrawingServiceParam[]>();
    remove$: Observable<ISheetDrawingServiceParam[]> = this._remove$.asObservable();

    private _update$ = new Subject<ISheetDrawingServiceUpdateParam[]>();
    update$: Observable<ISheetDrawingServiceUpdateParam[]> = this._update$.asObservable();

    private _focus$ = new Subject<ISheetDrawingServiceParam[]>();
    focus$: Observable<ISheetDrawingServiceParam[]> = this._focus$.asObservable();


    getDrawingMap(unitId: string, subUnitId: string, drawingType: DrawingType): Nullable<IISheetDrawingMapItem> {
        return this._drawingMap[drawingType]?.[unitId]?.[subUnitId];
    }

    getDrawingItem(param: IDrawingSearch): Nullable<ISheetDrawingServiceParam> {
        const { unitId, subUnitId, drawingId, drawingType } = param;
        return this._drawingMap[drawingType]?.[unitId]?.[subUnitId]?.[drawingId];
    }

    addDrawing(param: ISheetDrawingServiceParam): void {
        this._addDrawing(param);
        this._add$.next([param]);
    }

    batchAddDrawing(params: ISheetDrawingServiceParam[]): void {
        params.forEach((param) => {
            this._addDrawing(param);
        });
        this._add$.next(params);
    }

    removeDrawing(param: IDrawingSearch): Nullable<ISheetDrawingServiceParam> {
        const deleteItem = this._removeDrawing(param);

        deleteItem && this._remove$.next([deleteItem]);

        return deleteItem;
    }

    batchRemoveDrawing(params: IDrawingSearch[]): ISheetDrawingServiceParam[] {
        const deleteItems: ISheetDrawingServiceParam[] = [];
        params.forEach((param) => {
            const deleteItem = this._removeDrawing(param);
            deleteItem && deleteItems.push(deleteItem);
        });
        return deleteItems;
    }

    updateDrawing(param: ISheetDrawingServiceUpdateParam): void {
        this._updateDrawing(param);
        this._update$.next([param]);
    }

    batchUpdateDrawing(params: ISheetDrawingServiceUpdateParam[]): void {
        params.forEach((param) => {
            this._updateDrawing(param);
        });
        this._update$.next(params);
    }

    focusDrawing(param: Nullable<IDrawingSearch>): void {
        if (param == null) {
            this._focusDrawings = [];
            this._focus$.next([]);
            return;
        }
        const { unitId, subUnitId, drawingId, drawingType } = param;
        const item = this._drawingMap[drawingType]?.[unitId]?.[subUnitId]?.[drawingId];
        if (item == null) {
            this._focusDrawings = [];
            this._focus$.next([]);
            return;
        }
        this._focusDrawings.push(item);
        this._focus$.next([item]);
    }

    getFocusDrawings() {
        return this._focusDrawings;
    }

    private _updateDrawing(param: ISheetDrawingServiceUpdateParam): void {
        const { unitId, subUnitId, drawingId, drawingType } = param;

        if (!this._drawingMap[drawingType]) {
            return;
        }

        const drawingMapUnit = this._drawingMap[drawingType];

        if (!drawingMapUnit[unitId] || !drawingMapUnit[unitId][subUnitId]) {
            return;
        }
        const item = drawingMapUnit[unitId][subUnitId][drawingId];
        if (item == null) {
            return;
        }
        drawingMapUnit[unitId][subUnitId][drawingId] = { ...item, ...param };
    }


    private _addDrawing(param: ISheetDrawingServiceParam): void {
        const { unitId, subUnitId, drawingId, drawingType } = param;
        if (!this._drawingMap[drawingType]) {
            this._drawingMap[drawingType] = {};
        }
        if (!this._drawingMap[drawingType][unitId]) {
            this._drawingMap[drawingType][unitId] = {};
        }
        if (!this._drawingMap[drawingType][unitId][subUnitId]) {
            this._drawingMap[drawingType][unitId][subUnitId] = {};
        }
        this._drawingMap[drawingType][unitId][subUnitId][drawingId] = param;
    }

    private _removeDrawing(param: IDrawingSearch): Nullable<ISheetDrawingServiceParam> {
        const { unitId, subUnitId, drawingId, drawingType } = param;
        if (!this._drawingMap[drawingType]) {
            return;
        }

        const drawingMapUnit = this._drawingMap[drawingType];

        if (!drawingMapUnit[unitId] || !drawingMapUnit[unitId][subUnitId]) {
            return;
        }
        const deleteItem = drawingMapUnit[unitId][subUnitId][drawingId];
        delete drawingMapUnit[unitId][subUnitId][drawingId];
        return deleteItem;
    }
}

export const ISheetDrawingService = createIdentifier<SheetDrawingService>('univer.formula.sheet-drawing.service');