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
import { SetSelectionProtection } from '../mutation/set-selection-protection.mutation';
import { SelectionProtectionRuleModel } from '../../model';
import type { ISetRangeProtectionParams } from './type';

export const SetRangeProtectionCommand: ICommand<ISetRangeProtectionParams> = {
    type: CommandType.COMMAND,
    id: 'sheets.command.set-range-protection',
    async handler(accessor, params) {
        if (!params) {
            return false;
        }
        const commandService = accessor.get(ICommandService);
        const selectionProtectionModel = accessor.get(SelectionProtectionRuleModel);
        const undoRedoService = accessor.get(IUndoRedoService);
        const { rule, permissionId, oldRule } = params;

        const { unitId, subUnitId, ranges, name, description } = rule;

        if (rule.id) {
            const redoMutationParam = {
                unitId,
                subUnitId,
                ruleId: rule.id,
                rule: {
                    ranges,
                    permissionId,
                    id: selectionProtectionModel.createRuleId(unitId, subUnitId),
                    name,
                    description,
                },
            };
            const result = await commandService.executeCommand(SetSelectionProtection.id, redoMutationParam);

            if (result) {
                const redoMutations = [{
                    id: SetSelectionProtection.id, params: redoMutationParam,
                }];
                const undoMutations = [{
                    id: SetSelectionProtection.id, params: {
                        unitId,
                        subUnitId,
                        ruleId: rule.id,
                        rule: oldRule,
                    },
                }];
                undoRedoService.pushUndoRedo({
                    unitID: unitId,
                    redoMutations,
                    undoMutations,
                });
            }
        }
        return true;
    },
};
