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
import { CommandType, createDefaultUser, IUniverInstanceService, UniverInstanceType, UserManagerService } from '@univerjs/core';
import { WorkbookPermissionService } from '@univerjs/sheets';

import type { IAccessor } from '@wendellhu/redi';

export enum UnitRole {
    Reader = 0,
    Editor = 1,
    Owner = 2,
    UNRECOGNIZED = -1,
}

interface IChangeUserParams {
    value: UnitRole.Reader | UnitRole.Owner;

}

export const ChangeUserCommand: ICommand = {
    id: 'debugger.operation.changeUser',
    type: CommandType.OPERATION,
    handler: async (accessor: IAccessor, params: IChangeUserParams) => {
        const userManagerService = accessor.get(UserManagerService);
        const univerInstanceService = accessor.get(IUniverInstanceService);
        userManagerService.currentUser = createDefaultUser(params.value);
        const workbook = univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const workbookPermissionService = accessor.get(WorkbookPermissionService);
        const unitId = workbook.getUnitId();
        if (params.value === UnitRole.Owner) {
            workbookPermissionService.setManageCollaboratorPermission(unitId, true);
        } else {
            workbookPermissionService.setManageCollaboratorPermission(unitId, false);
        }
        return true;
    },
};
