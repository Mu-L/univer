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

import { Inject } from '@wendellhu/redi';
import type { IPermissionPoint, Workbook } from '@univerjs/core';
import { Disposable, IPermissionService, IUniverInstanceService, LifecycleStages, OnLifecycle, UnitPermissionType, UniverInstanceType } from '@univerjs/core';
import { map, of } from 'rxjs';
import {
    WorkbookCommentPermission,
    WorkbookCopyPermission,
    WorkbookCreateSheetPermission,
    WorkbookDeleteSheetPermission,
    WorkbookDuplicatePermission,
    WorkbookEditablePermission,
    WorkbookExportPermission,
    WorkbookHideSheetPermission,
    WorkbookHistoryPermission,
    WorkbookManageCollaboratorPermission,
    WorkbookMoveSheetPermission,
    WorkbookPrintPermission,
    WorkbookRenameSheetPermission,
    WorkbookSharePermission,
    WorkbookViewPermission,
} from '../permission-point';
import type { GetWorkbookPermissionFunc, GetWorkbookPermissionFunc$, SetWorkbookPermissionFunc } from '../type';
import { getAllWorkbookPermissionPoint } from './utils';

@OnLifecycle(LifecycleStages.Starting, WorkbookPermissionService)
export class WorkbookPermissionService extends Disposable {
    getEditPermission$: GetWorkbookPermissionFunc$;
    getEditPermission: GetWorkbookPermissionFunc;
    setEditPermission: SetWorkbookPermissionFunc;

    getPrintPermission$: GetWorkbookPermissionFunc$;
    getPrintPermission: GetWorkbookPermissionFunc;
    setPrintPermission: SetWorkbookPermissionFunc;

    getDuplicatePermission$: GetWorkbookPermissionFunc$;
    getDuplicatePermission: GetWorkbookPermissionFunc;
    setDuplicatePermission: SetWorkbookPermissionFunc;

    getExportPermission$: GetWorkbookPermissionFunc$;
    getExportPermission: GetWorkbookPermissionFunc;
    setExportPermission: SetWorkbookPermissionFunc;

    getMoveSheetPermission$: GetWorkbookPermissionFunc$;
    getMoveSheetPermission: GetWorkbookPermissionFunc;
    setMoveSheetPermission: SetWorkbookPermissionFunc;

    getDeleteSheetPermission$: GetWorkbookPermissionFunc$;
    getDeleteSheetPermission: GetWorkbookPermissionFunc;
    setDeleteSheetPermission: SetWorkbookPermissionFunc;

    getHideSheetPermission$: GetWorkbookPermissionFunc$;
    getHideSheetPermission: GetWorkbookPermissionFunc;
    setHideSheetPermission: SetWorkbookPermissionFunc;

    getRenameSheetPermission$: GetWorkbookPermissionFunc$;
    getRenameSheetPermission: GetWorkbookPermissionFunc;
    setRenameSheetPermission: SetWorkbookPermissionFunc;

    getCreateSheetPermission$: GetWorkbookPermissionFunc$;
    getCreateSheetPermission: GetWorkbookPermissionFunc;
    setCreateSheetPermission: SetWorkbookPermissionFunc;

    getHistoryPermission$: GetWorkbookPermissionFunc$;
    getHistoryPermission: GetWorkbookPermissionFunc;
    setHistoryPermission: SetWorkbookPermissionFunc;

    getViewPermission$: GetWorkbookPermissionFunc$;
    getViewPermission: GetWorkbookPermissionFunc;
    setViewPermission: SetWorkbookPermissionFunc;

    getSharePermission$: GetWorkbookPermissionFunc$;
    getSharePermission: GetWorkbookPermissionFunc;
    setSharePermission: SetWorkbookPermissionFunc;

    getCommentPermission$: GetWorkbookPermissionFunc$;
    getCommentPermission: GetWorkbookPermissionFunc;
    setCommentPermission: SetWorkbookPermissionFunc;

    getCopyPermission$: GetWorkbookPermissionFunc$;
    getCopyPermission: GetWorkbookPermissionFunc;
    setCopyPermission: SetWorkbookPermissionFunc;

    getProtectSheetPermission$: GetWorkbookPermissionFunc$;
    getProtectSheetPermission: GetWorkbookPermissionFunc;
    setProtectSheetPermission: SetWorkbookPermissionFunc;

    getCopySheetPermission$: GetWorkbookPermissionFunc$;
    getCopySheetPermission: GetWorkbookPermissionFunc;
    setCopySheetPermission: SetWorkbookPermissionFunc;

    getCollaboratorPermission$: GetWorkbookPermissionFunc$;
    getCollaboratorPermission: GetWorkbookPermissionFunc;
    setCollaboratorPermission: SetWorkbookPermissionFunc;

    getManageCollaboratorPermission$: GetWorkbookPermissionFunc$;
    getManageCollaboratorPermission: GetWorkbookPermissionFunc;
    setManageCollaboratorPermission: SetWorkbookPermissionFunc;

    constructor(
        @Inject(IPermissionService) private _permissionService: IPermissionService,
        @Inject(IUniverInstanceService) private _univerInstanceService: IUniverInstanceService
    ) {
        super();
        this._init();
        this._initializePermissions();
    }

    private _init() {
        const handleWorkbook = (workbook: Workbook) => {
            const unitId = workbook.getUnitId();
            getAllWorkbookPermissionPoint().forEach((F) => {
                const instance = new F(unitId);
                this._permissionService.addPermissionPoint(instance);
            });
        };

        this._univerInstanceService.getAllUnitsForType<Workbook>(UniverInstanceType.UNIVER_SHEET).forEach((workbook) => {
            handleWorkbook(workbook);
        });

        this.disposeWithMe(this._univerInstanceService.getTypeOfUnitAdded$<Workbook>(UniverInstanceType.UNIVER_SHEET).subscribe((workbook) => {
            handleWorkbook(workbook);
        }));

        this.disposeWithMe(this._univerInstanceService.getTypeOfUnitDisposed$<Workbook>(UniverInstanceType.UNIVER_SHEET).subscribe((workbook) => {
            const unitId = workbook.getUnitId();
            getAllWorkbookPermissionPoint().forEach((F) => {
                const instance = new F(unitId);
                this._permissionService.deletePermissionPoint(instance.id);
            });
        }));
    }

    private _createPermissionMethods<T extends IPermissionPoint>(PermissionClass: new (unitId: string) => T) {
        return {
            get$: (unitId: string) => {
                const permissionInstance = new PermissionClass(unitId);
                const permission = this._permissionService.getPermissionPoint(permissionInstance.id);
                if (!permission) {
                    return of(false);
                }
                return this._permissionService.composePermission$([permissionInstance.id]).pipe(map((list) => {
                    return list.every((item) => item.value === true);
                }));
            },
            get: (unitId: string) => {
                if (!unitId) return false;
                const permissionInstance = new PermissionClass(unitId);
                const permission = this._permissionService.getPermissionPoint(permissionInstance.id);
                return permission?.value ?? false;
            },
            set: (unitId: string, value: boolean) => {
                const permissionInstance = new PermissionClass(unitId);
                this._permissionService.updatePermissionPoint(permissionInstance.id, value);
            },
        };
    }

    private _initializePermissions() {
        const permissions = [
            { type: UnitPermissionType.Edit, class: WorkbookEditablePermission },
            { type: UnitPermissionType.Print, class: WorkbookPrintPermission },
            { type: UnitPermissionType.Duplicate, class: WorkbookDuplicatePermission },
            { type: UnitPermissionType.Export, class: WorkbookExportPermission },
            { type: UnitPermissionType.MoveSheet, class: WorkbookMoveSheetPermission },
            { type: UnitPermissionType.DeleteSheet, class: WorkbookDeleteSheetPermission },
            { type: UnitPermissionType.HideSheet, class: WorkbookHideSheetPermission },
            { type: UnitPermissionType.RenameSheet, class: WorkbookRenameSheetPermission },
            { type: UnitPermissionType.CreateSheet, class: WorkbookCreateSheetPermission },
            { type: UnitPermissionType.History, class: WorkbookHistoryPermission },
            { type: UnitPermissionType.View, class: WorkbookViewPermission },
            { type: UnitPermissionType.Share, class: WorkbookSharePermission },
            { type: UnitPermissionType.Comment, class: WorkbookCommentPermission },
            { type: UnitPermissionType.Copy, class: WorkbookCopyPermission },
            { type: UnitPermissionType.CopySheet, class: WorkbookCopyPermission },
            { type: UnitPermissionType.ManageCollaborator, class: WorkbookManageCollaboratorPermission },
        ];

        permissions.forEach(({ type, class: PermissionClass }) => {
            const { get$, get, set } = this._createPermissionMethods(PermissionClass);
            this[`get${type}Permission$`] = get$;
            this[`get${type}Permission`] = get;
            this[`set${type}Permission`] = set;
        });
    }
}
