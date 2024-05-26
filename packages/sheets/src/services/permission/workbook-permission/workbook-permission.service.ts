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
import { Disposable, IPermissionService, IUniverInstanceService, LifecycleStages, OnLifecycle, UniverInstanceType } from '@univerjs/core';
import { map, of } from 'rxjs';
import { UnitAction } from '@univerjs/protocol';
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
import type { IWorkbookPermissionServiceMethods } from '../type';
import { getAllWorkbookPermissionPoint } from './utils';

@OnLifecycle(LifecycleStages.Starting, WorkbookPermissionService)
export class WorkbookPermissionService extends Disposable implements IWorkbookPermissionServiceMethods {
    [key: string]: any;

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
            { type: UnitAction.Edit, class: WorkbookEditablePermission },
            { type: UnitAction.Print, class: WorkbookPrintPermission },
            { type: UnitAction.Duplicate, class: WorkbookDuplicatePermission },
            { type: UnitAction.Export, class: WorkbookExportPermission },
            { type: UnitAction.MoveSheet, class: WorkbookMoveSheetPermission },
            { type: UnitAction.DeleteSheet, class: WorkbookDeleteSheetPermission },
            { type: UnitAction.HideSheet, class: WorkbookHideSheetPermission },
            { type: UnitAction.RenameSheet, class: WorkbookRenameSheetPermission },
            { type: UnitAction.CreateSheet, class: WorkbookCreateSheetPermission },
            { type: UnitAction.History, class: WorkbookHistoryPermission },
            { type: UnitAction.View, class: WorkbookViewPermission },
            { type: UnitAction.Share, class: WorkbookSharePermission },
            { type: UnitAction.Comment, class: WorkbookCommentPermission },
            { type: UnitAction.Copy, class: WorkbookCopyPermission },
            { type: UnitAction.CopySheet, class: WorkbookCopyPermission },
            { type: UnitAction.ManageCollaborator, class: WorkbookManageCollaboratorPermission },
        ];

        permissions.forEach(({ type, class: PermissionClass }) => {
            const { get$, get, set } = this._createPermissionMethods(PermissionClass);
            this[`get${type}Permission$`] = get$;
            this[`get${type}Permission`] = get;
            this[`set${type}Permission`] = set;
        });
    }
}
