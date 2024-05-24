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

import type { IAccessor } from '@wendellhu/redi';
import type { IMutation, IMutationInfo } from '@univerjs/core';
import { CommandType } from '@univerjs/core';

import { SelectionProtectionRuleModel } from '../../model/selection-protection-rule.model';
import type { ISelectionProtectionRule } from '../../model/type';
import type { IAddSelectionProtectionParams } from './add-selection-protection.mutation';
import { AddSelectionProtection } from './add-selection-protection.mutation';

export interface IDeleteSelectionProtectionParams {
    ruleIds: string[];
    unitId: string;
    subUnitId: string;
};
export const FactoryDeleteSelectionProtection = (accessor: IAccessor, param: IDeleteSelectionProtectionParams) => {
    const selectionProtectionRuleModel = accessor.get(SelectionProtectionRuleModel);
    const rules = param.ruleIds.map((id) => selectionProtectionRuleModel.getRule(param.unitId, param.subUnitId, id)).filter((rule) => !!rule) as ISelectionProtectionRule[];
    const result: IMutationInfo<Omit<IAddSelectionProtectionParams, 'name'>> = { id: AddSelectionProtection.id, params: { subUnitId: param.subUnitId, unitId: param.unitId, rules } };
    return result;
};
export const DeleteSelectionProtection: IMutation<IDeleteSelectionProtectionParams> = {
    id: 'sheet.mutation.delete.selection-protection',
    type: CommandType.MUTATION,
    handler: (accessor, params) => {
        const { unitId, subUnitId, ruleIds } = params;
        const selectionProtectionRuleModel = accessor.get(SelectionProtectionRuleModel);
        ruleIds.forEach((id) => {
            selectionProtectionRuleModel.deleteRule(unitId, subUnitId, id);
        });
        return true;
    },
};
