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

import type { ICommand, Workbook } from '@univerjs/core';
import { CommandType, IUniverInstanceService, UniverInstanceType } from '@univerjs/core';
import { WorkbookPermissionService, WorksheetPermissionService } from '@univerjs/sheets';
import type { IAccessor } from '@wendellhu/redi';

export interface ISetEditableCommandParams {
    value: 'sheet' | 'univer';
}

export const SetEditable: ICommand = {
    id: 'debugger.operation.set.editable',
    type: CommandType.OPERATION,
    handler: (accessor: IAccessor, params: ISetEditableCommandParams) => {
        const univerInstanceService = accessor.get(IUniverInstanceService);
        const workbook = univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET);
        const worksheet = workbook?.getActiveSheet();
        if (!workbook || !worksheet) {
            return false;
        }
        if (params.value === 'sheet') {
            const workSheetPermissionService = accessor.get(WorksheetPermissionService);
            const editable = workSheetPermissionService.getSetCellValuePermission({
                unitId: workbook.getUnitId(),
                subUnitId: worksheet.getSheetId(),
            });
            workSheetPermissionService.setSetCellValuePermission(!editable);
        } else {
            const workbookPermissionService = accessor.get(WorkbookPermissionService);
            const unitId = workbook!.getUnitId();
            const editable = workbookPermissionService.getEditPermission(unitId);
            workbookPermissionService.setEditPermission(unitId, !editable);
        }
        return true;
    },
};
