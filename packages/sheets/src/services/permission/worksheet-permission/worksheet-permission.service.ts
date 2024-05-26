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

import type { IPermissionParam, IPermissionPoint, Workbook, Worksheet } from '@univerjs/core';
import { IPermissionService, IResourceManagerService, IUniverInstanceService, LifecycleStages, OnLifecycle, RxDisposable, UniverInstanceType } from '@univerjs/core';
import { Inject, Injector } from '@wendellhu/redi';
import { map, takeUntil } from 'rxjs/operators';

// import type { UnitAction, UnitObject } from '@univerjs/protocol';
import { UnitAction, UniverType } from '@univerjs/protocol';
import type { ISheetFontRenderExtension } from '@univerjs/engine-render';
import {
    WorkbookCopyPermission,
    WorkbookEditablePermission,
    WorkbookManageCollaboratorPermission,
    WorkbookViewPermission,
    WorksheetCopyPermission,
    WorksheetDeleteColumnPermission,
    WorksheetDeleteRowPermission,
    WorksheetEditExtraObjectPermission,
    WorksheetEditPermission,
    WorksheetFilterPermission,
    WorksheetInsertColumnPermission,
    WorksheetInsertHyperlinkPermission,
    WorksheetInsertRowPermission,
    WorksheetManageCollaboratorPermission,
    WorksheetPivotTablePermission,
    WorksheetSelectProtectedCellsPermission,
    WorksheetSelectUnProtectedCellsPermission,
    WorksheetSetCellStylePermission,
    WorksheetSetCellValuePermission,
    WorksheetSetColumnStylePermission,
    WorksheetSetRowStylePermission,
    WorksheetSortPermission,
    WorksheetViewPermission,
} from '../permission-point';
import type { IObjectModel, IObjectPointModel, IWorksheetPermissionServiceMethods } from '../type';
import { SheetInterceptorService } from '../../sheet-interceptor/sheet-interceptor.service';
import { INTERCEPTOR_POINT } from '../../sheet-interceptor/interceptor-const';
import { WorksheetProtectionRuleModel } from './worksheet-permission-rule.model';
import { getAllWorksheetPermissionPoint, getAllWorksheetPermissionPointByPointPanel } from './utils';
import type { IWorksheetProtectionRenderCellData } from './type';
import { WorksheetProtectionPointModel } from './worksheet-permission-point.model';

export const RULE_MODEL_PLUGIN_NAME = 'SHEET_WORKSHEET_PROTECTION_PLUGIN';
export const POINT_MODEL_PLUGIN_NAME = 'SHEET_WORKSHEET_PROTECTION_POINT_PLUGIN';

@OnLifecycle(LifecycleStages.Starting, WorksheetPermissionService)
export class WorksheetPermissionService extends RxDisposable implements IWorksheetPermissionServiceMethods {
    [key: string]: any;

    constructor(
        @Inject(IPermissionService) private _permissionService: IPermissionService,
        @Inject(IUniverInstanceService) private _univerInstanceService: IUniverInstanceService,
        @Inject(Injector) readonly _injector: Injector,
        @Inject(WorksheetProtectionRuleModel) private _worksheetProtectionRuleModel: WorksheetProtectionRuleModel,
        @Inject(WorksheetProtectionPointModel) private _worksheetProtectionPointRuleModel: WorksheetProtectionPointModel,
        @Inject(IResourceManagerService) private _resourceManagerService: IResourceManagerService,
        @Inject(SheetInterceptorService) private _sheetInterceptorService: SheetInterceptorService
    ) {
        super();
        this._init();
        this._initializePermissions();
        this._initRuleChange();
        this._initRuleSnapshot();
        this._initPointSnapshot();
        this._initViewModelInterceptor();
    }

    private _init() {
        const handleWorkbook = (workbook: Workbook) => {
            const unitId = workbook.getUnitId();
            const handleWorksheet = (worksheet: Worksheet) => {
                const subUnitId = worksheet.getSheetId();
                [...getAllWorksheetPermissionPoint(), ...getAllWorksheetPermissionPointByPointPanel()].forEach((F) => {
                    const instance = new F(unitId, subUnitId);
                    this._permissionService.addPermissionPoint(instance);
                });
            };
            workbook.getSheets().forEach((worksheet) => {
                handleWorksheet(worksheet);
            });
            workbook.sheetCreated$.subscribe((worksheet) => {
                handleWorksheet(worksheet);
            });
            workbook.sheetDisposed$.subscribe((worksheet) => {
                const subUnitId = worksheet.getSheetId();
                [...getAllWorksheetPermissionPoint(), ...getAllWorksheetPermissionPointByPointPanel()].forEach((F) => {
                    const instance = new F(unitId, subUnitId);
                    this._permissionService.deletePermissionPoint(instance.id);
                });
            });
        };

        this._univerInstanceService.getAllUnitsForType<Workbook>(UniverInstanceType.UNIVER_SHEET).forEach((workbook) => {
            handleWorkbook(workbook);
        });

        this._univerInstanceService.getTypeOfUnitAdded$<Workbook>(UniverInstanceType.UNIVER_SHEET).pipe(takeUntil(this.dispose$)).subscribe(handleWorkbook);

        this._univerInstanceService.getTypeOfUnitDisposed$<Workbook>(UniverInstanceType.UNIVER_SHEET).pipe(takeUntil(this.dispose$)).subscribe((workbook) => {
            workbook.getSheets().forEach((worksheet) => {
                const unitId = workbook.getUnitId();
                const subUnitId = worksheet.getSheetId();
                getAllWorksheetPermissionPoint().forEach((F) => {
                    const instance = new F(unitId, subUnitId);
                    this._permissionService.deletePermissionPoint(instance.id);
                });
            });
        });
    }

    private _createPermissionMethods<T extends IPermissionPoint>(PermissionClassGroup: { WorkbookPermissionClass: new (unitId: string) => T; WorksheetPermissionClass: new (unitId: string, subUnitId: string) => T }) {
        return {
            get$: (permissionParams: IPermissionParam) => {
                const { unitId, subUnitId } = permissionParams;
                const { WorkbookPermissionClass, WorksheetPermissionClass } = PermissionClassGroup;
                const unitPermissionInstance = new WorkbookPermissionClass(unitId);
                const subUnitPermissionInstance = new WorksheetPermissionClass(unitId, subUnitId);
                const workbookPermission = this._permissionService.getPermissionPoint$(unitPermissionInstance.id);
                const sheetPermission = this._permissionService.getPermissionPoint$(subUnitPermissionInstance.id);
                if (!sheetPermission || !workbookPermission) {
                    throw (new Error('Permission initialization error.'));
                }
                return this._permissionService.composePermission$([unitPermissionInstance.id, subUnitPermissionInstance.id]).pipe(map((list) => {
                    return list.every((item) => item.value === true);
                }));
            },
            get: (permissionParams: IPermissionParam) => {
                const { unitId, subUnitId } = permissionParams;
                const { WorkbookPermissionClass, WorksheetPermissionClass } = PermissionClassGroup;
                const workbookPermissionInstance = new WorkbookPermissionClass(unitId);
                const worksheetPermissionInstance = new WorksheetPermissionClass(unitId, subUnitId);
                const workbookPermission = this._permissionService.getPermissionPoint(workbookPermissionInstance.id);
                const sheetPermission = this._permissionService.getPermissionPoint(worksheetPermissionInstance.id);
                if (!sheetPermission || !workbookPermission) {
                    return false;
                }
                return workbookPermission.value && sheetPermission.value;
            },
            set: (value: boolean, unitId?: string, subUnitId?: string) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                if (!workbook) return;
                const _unitId = unitId || workbook.getUnitId();
                const sheet = workbook.getActiveSheet();
                const _subUnitId = subUnitId || sheet.getSheetId();
                const { WorksheetPermissionClass } = PermissionClassGroup;
                const sheetPermission = new WorksheetPermissionClass(_unitId, _subUnitId);
                this._permissionService.updatePermissionPoint(sheetPermission.id, value);
            },
        };
    }

    // eslint-disable-next-line max-lines-per-function
    private _initializePermissions() {
        const permissions = [
            {
                type: UnitAction.Copy,
                classGroup: {
                    WorkbookPermissionClass: WorkbookCopyPermission,
                    WorksheetPermissionClass: WorksheetCopyPermission,
                },
            },
            {
                type: UnitAction.SelectProtectedCells,
                classGroup: {
                    WorkbookPermissionClass: WorkbookViewPermission,
                    WorksheetPermissionClass: WorksheetSelectProtectedCellsPermission,
                },
            },
            {
                type: UnitAction.SelectUnProtectedCells,
                classGroup: {
                    WorkbookPermissionClass: WorkbookViewPermission,
                    WorksheetPermissionClass: WorksheetSelectUnProtectedCellsPermission,
                },
            },
            {
                type: UnitAction.SetCellStyle,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetSetCellStylePermission,
                },
            },
            {
                type: UnitAction.SetCellValue,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetSetCellValuePermission,
                },
            },
            {
                type: UnitAction.View,
                classGroup: {
                    WorkbookPermissionClass: WorkbookViewPermission,
                    WorksheetPermissionClass: WorksheetViewPermission,
                },
            },
            {
                type: UnitAction.SetRowStyle,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetSetRowStylePermission,
                },
            },
            {
                type: UnitAction.SetColumnStyle,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetSetColumnStylePermission,
                },
            },
            {
                type: UnitAction.InsertRow,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetInsertRowPermission,
                },
            },
            {
                type: UnitAction.InsertColumn,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetInsertColumnPermission,
                },
            },
            {
                type: UnitAction.InsertHyperlink,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetInsertHyperlinkPermission,
                },
            },
            {
                type: UnitAction.DeleteRow,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetDeleteRowPermission,
                },
            },
            {
                type: UnitAction.DeleteColumn,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetDeleteColumnPermission,
                },
            },
            {
                type: UnitAction.Sort,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetSortPermission,
                },
            },
            {
                type: UnitAction.Filter,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetFilterPermission,
                },
            },
            {
                type: UnitAction.PivotTable,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetPivotTablePermission,
                },
            },
            {
                type: UnitAction.EditExtraObject,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetEditExtraObjectPermission,
                },
            },
            {
                type: UnitAction.ManageCollaborator,
                classGroup: {
                    WorkbookPermissionClass: WorkbookManageCollaboratorPermission,
                    WorksheetPermissionClass: WorksheetManageCollaboratorPermission,
                },
            },
            {
                type: UnitAction.Edit,
                classGroup: {
                    WorkbookPermissionClass: WorkbookEditablePermission,
                    WorksheetPermissionClass: WorksheetEditPermission,
                },
            },
        ];

        permissions.forEach(({ type, classGroup }) => {
            const { get$, get, set } = this._createPermissionMethods(classGroup as { WorkbookPermissionClass: new (unitId: string) => IPermissionPoint; WorksheetPermissionClass: new (unitId: string, subUnitId: string) => IPermissionPoint });
            this[`get${type}Permission$`] = get$;
            this[`get${type}Permission`] = get;
            this[`set${type}Permission`] = set;
        });
    }

    private _initRuleChange() {
        this.disposeWithMe(
            this._worksheetProtectionRuleModel.ruleChange$.subscribe((info) => {
                switch (info.type) {
                    case 'add': {
                        getAllWorksheetPermissionPoint().forEach((F) => {
                            const instance = new F(info.unitId, info.subUnitId);
                            this._permissionService.addPermissionPoint(instance);
                        });
                        break;
                    }
                    case 'delete': {
                        getAllWorksheetPermissionPoint().forEach((F) => {
                            const instance = new F(info.unitId, info.subUnitId);
                            this._permissionService.deletePermissionPoint(instance.id);
                        });
                        break;
                    }
                    case 'set': {
                        getAllWorksheetPermissionPoint().forEach((F) => {
                            const instance = new F(info.unitId, info.subUnitId);
                            this._permissionService.updatePermissionPoint(instance.id, info.rule);
                        });
                        break;
                    }
                }
            }));
    }

    private _initRuleSnapshot() {
        const toJson = () => {
            const object = this._worksheetProtectionRuleModel.toObject();
            return JSON.stringify(object);
        };
        const parseJson = (json: string): IObjectModel => {
            if (!json) {
                return {};
            }
            try {
                return JSON.parse(json);
            } catch (err) {
                return {};
            }
        };
        this.disposeWithMe(
            this._resourceManagerService.registerPluginResource({
                toJson, parseJson,
                pluginName: RULE_MODEL_PLUGIN_NAME,
                businesses: [UniverType.UNIVER_SHEET],
                onLoad: (unitId, resources) => {
                    this._worksheetProtectionRuleModel.fromObject(resources);
                    Object.keys(resources).forEach((subUnitId) => {
                        getAllWorksheetPermissionPoint().forEach((F) => {
                            const instance = new F(unitId, subUnitId);
                            this._permissionService.addPermissionPoint(instance);
                        });
                    });
                    this._worksheetProtectionRuleModel.changeRuleInitState(true);
                },
                onUnLoad: () => {
                    this._worksheetProtectionRuleModel.deleteUnitModel();
                },
            })
        );
    }

    private _initPointSnapshot() {
        const toJson = () => {
            const object = this._worksheetProtectionPointRuleModel.toObject();
            return JSON.stringify(object);
        };
        const parseJson = (json: string): IObjectPointModel => {
            if (!json) {
                return {};
            }
            try {
                return JSON.parse(json);
            } catch (err) {
                return {};
            }
        };
        this.disposeWithMe(
            this._resourceManagerService.registerPluginResource({
                toJson, parseJson,
                pluginName: POINT_MODEL_PLUGIN_NAME,
                businesses: [UniverType.UNIVER_SHEET],
                onLoad: (unitId, resources) => {
                    this._worksheetProtectionPointRuleModel.fromObject(resources);
                    Object.keys(resources).forEach((subUnitId) => {
                        getAllWorksheetPermissionPointByPointPanel().forEach((F) => {
                            const instance = new F(unitId, subUnitId);
                            this._permissionService.addPermissionPoint(instance);
                        });
                    });
                },
                onUnLoad: () => {
                    this._worksheetProtectionPointRuleModel.deleteUnitModel();
                },
            })
        );
    }

    private _initViewModelInterceptor() {
        this.disposeWithMe(this._sheetInterceptorService.intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
            handler: (cell = {}, context, next) => {
                const { unitId, subUnitId } = context;
                const worksheetRule = this._worksheetProtectionRuleModel.getRule(unitId, subUnitId);
                if (worksheetRule?.permissionId && worksheetRule.name) {
                    const selectionProtection = [{
                        View: this._permissionService.getPermissionPoint(new WorksheetViewPermission(unitId, subUnitId).id)?.value ?? false,
                        Edit: this._permissionService.getPermissionPoint(new WorksheetEditPermission(unitId, subUnitId).id)?.value ?? false,
                    }];
                    const isSkipFontRender = !selectionProtection[0].View;
                    const _cellData: IWorksheetProtectionRenderCellData & ISheetFontRenderExtension = { ...cell, hasWorksheetRule: true, selectionProtection };
                    if (isSkipFontRender) {
                        if (!_cellData.fontRenderExtension) {
                            _cellData.fontRenderExtension = {};
                        }
                        _cellData.fontRenderExtension.isSkip = isSkipFontRender;
                    }
                    return next(_cellData);
                }
                return next(cell);
            },
        }
        ));
    }
}
