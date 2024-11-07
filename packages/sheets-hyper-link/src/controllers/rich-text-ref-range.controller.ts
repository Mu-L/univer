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

import type { IDisposable, IDocumentData, IRange, Nullable, Workbook } from '@univerjs/core';
import type { ISetRangeValuesMutationParams } from '@univerjs/sheets';
import { CustomRangeType, Disposable, DisposableCollection, ICommandService, Inject, isValidRange, IUniverInstanceService, ObjectMatrix, UniverInstanceType } from '@univerjs/core';
import { deserializeRangeWithSheet, serializeRange } from '@univerjs/engine-formula';
import { RefRangeService, SetRangeValuesMutation } from '@univerjs/sheets';
import { ERROR_RANGE } from '../types/const';

export class SheetsHyperLinkRichTextRefRangeController extends Disposable {
    private _refRangeMap: Map<string, Map<string, ObjectMatrix<IDisposable>>> = new Map();
    constructor(
        @ICommandService private readonly _commandService: ICommandService,
        @IUniverInstanceService private readonly _univerInstanceService: IUniverInstanceService,
        @Inject(RefRangeService) private readonly _refRangeService: RefRangeService
    ) {
        super();
        this._initWorkbookLoad();
        this._initWorkbookUnload();
        this._initSetRangesListener();
    }

    private _enusreMap(unitId: string, subUnitId: string) {
        let unitMap = this._refRangeMap.get(unitId);
        if (!unitMap) {
            unitMap = new Map();
            this._refRangeMap.set(unitId, unitMap);
        }
        let subUnitMap = unitMap.get(subUnitId);
        if (!subUnitMap) {
            subUnitMap = new ObjectMatrix();
            unitMap.set(subUnitId, subUnitMap);
        }
        return subUnitMap;
    }

    private _isLegalRangeUrl(unitId: string, payload: string): Nullable<IRange> {
        const workbook = this._univerInstanceService.getUnit<Workbook>(unitId, UniverInstanceType.UNIVER_SHEET);
        if (!workbook) {
            return null;
        }
        if (payload && payload.startsWith('#')) {
            const search = new URLSearchParams(payload.slice(1));

            // range, gid, rangeid
            const searchObj = {
                gid: search.get('gid') ?? '',
                range: search.get('range') ?? '',
                rangeid: search.get('rangeid') ?? '',
            };

            if (searchObj.range && searchObj.gid) {
                const subUnitId = searchObj.gid;
                const worksheet = workbook.getSheetBySheetId(subUnitId);
                if (!worksheet) {
                    return null;
                }
                const range = deserializeRangeWithSheet(searchObj.range).range;
                if (isValidRange(range, worksheet) && searchObj.range !== ERROR_RANGE) {
                    return range;
                }
            }
        }

        return null;
    }

    private _registerRange(unitId: string, subUnitId: string, row: number, col: number, p: IDocumentData) {
        const map = this._enusreMap(unitId, subUnitId);

        if (p.body?.customRanges?.some((customRange) => customRange.rangeType === CustomRangeType.HYPERLINK && this._isLegalRangeUrl(unitId, customRange.properties?.url))) {
            const disposableCollection = new DisposableCollection();
            p.body?.customRanges?.forEach((customRange) => {
                if (customRange.rangeType === CustomRangeType.HYPERLINK) {
                    const payload = customRange.properties?.url;
                    const range = this._isLegalRangeUrl(unitId, payload);
                    if (range) {
                        disposableCollection.add(this._refRangeService.watchRange(unitId, subUnitId, range, (before, after) => {
                            customRange.properties!.url = `#gid=${subUnitId}&range=${after ? serializeRange(after) : ERROR_RANGE}`;
                        }));
                    }
                }
            });

            map.setValue(row, col, disposableCollection);
        }
    }

    private _initWorkbookLoad() {
        const handleWorkbook = (workbook: Workbook) => {
            const unitId = workbook.getUnitId();
            workbook.getSheets().forEach((sheet) => {
                const subUnitId = sheet.getSheetId();
                const map = this._enusreMap(unitId, subUnitId);
                sheet.getCellMatrix().forValue((row, col, cell) => {
                    const dispose = map.getValue(row, col);
                    if (dispose) {
                        dispose.dispose();
                    }

                    if (cell && cell.p) {
                        this._registerRange(unitId, subUnitId, row, col, cell.p);
                    }
                });
            });
        };

        this._univerInstanceService.getAllUnitsForType<Workbook>(UniverInstanceType.UNIVER_SHEET).forEach((workbook) => {
            handleWorkbook(workbook);
        });
        this.disposeWithMe(
            this._univerInstanceService.unitAdded$.subscribe((unit) => {
                if (unit.type === UniverInstanceType.UNIVER_SHEET) {
                    const workbook = unit as Workbook;
                    handleWorkbook(workbook);
                }
            })
        );
    }

    private _initWorkbookUnload() {
        this._univerInstanceService.unitDisposed$.subscribe((unit) => {
            if (unit.type === UniverInstanceType.UNIVER_SHEET) {
                const workbook = unit as Workbook;
                const unitId = workbook.getUnitId();
                workbook.getSheets().forEach((sheet) => {
                    const subUnitId = sheet.getSheetId();
                    const map = this._enusreMap(unitId, subUnitId);
                    map.forValue((row, col, dispose) => {
                        if (dispose) {
                            dispose.dispose();
                        }
                    });
                });
                this._refRangeMap.delete(unitId);
            }
        });
    }

    private _initSetRangesListener() {
        this.disposeWithMe(
            this._commandService.onCommandExecuted((commandInfo) => {
                if (commandInfo.id === SetRangeValuesMutation.id) {
                    const params = commandInfo.params as ISetRangeValuesMutationParams;
                    const { unitId, subUnitId, cellValue } = params;
                    const map = this._enusreMap(unitId, subUnitId);
                    if (cellValue) {
                        new ObjectMatrix(cellValue).forValue((row, col, cell) => {
                            const dispose = map.getValue(row, col);
                            if (dispose) {
                                dispose.dispose();
                            }
                            if (cell && cell.p) {
                                this._registerRange(unitId, subUnitId, row, col, cell.p);
                            }
                        });
                    }
                }
            })
        );
    }
}