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

import type { ICommand } from '@univerjs/core';
import { CommandType, ICommandService, IUndoRedoService } from '@univerjs/core';
import { DeleteSelectionProtection } from '../mutation/delete-selection-protection.mutation';
import { AddSelectionProtection } from '../mutation/add-selection-protection.mutation';
import type { IDeleteRangeProtectionParams } from './type';

export const DeleteRangeProtectionCommand: ICommand<IDeleteRangeProtectionParams> = {
    type: CommandType.COMMAND,
    id: 'sheets.command.delete-range-protection',
    async handler(accessor, params) {
        if (!params) {
            return false;
        }
        const commandService = accessor.get(ICommandService);
        const undoRedoService = accessor.get(IUndoRedoService);
        const { unitId, subUnitId, rule } = params;

        const redoMutationParam = {
            unitId,
            subUnitId,
            ruleIds: [rule.id],
        };
        const result = await commandService.executeCommand(DeleteSelectionProtection.id, redoMutationParam);

        if (result) {
            undoRedoService.pushUndoRedo({
                unitID: unitId,
                redoMutations: [{ id: DeleteSelectionProtection.id, params: redoMutationParam }],
                undoMutations: [{ id: AddSelectionProtection.id, params: { unitId, subUnitId, rules: [rule] } }],
            });
        }

        return true;
    },
};
