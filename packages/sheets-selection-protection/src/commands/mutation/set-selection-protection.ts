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

import type { ISelectionProtectionRule } from '../../model/type';
import { SelectionProtectionRuleModel } from '../../model/selection-protection-rule.model';

export interface IAddSelectionProtectionParams {
    rule: ISelectionProtectionRule;
    unitId: string;
    subUnitId: string;
    ruleId: string;
};

export type ISetSelectionProtectionParams = IAddSelectionProtectionParams;

export const SetSelectionProtection: IMutation<ISetSelectionProtectionParams> = {
    id: 'sheet.mutation.set-selection-protection',
    type: CommandType.MUTATION,
    handler: (accessor, params) => {
        const { unitId, subUnitId, rule, ruleId } = params;
        const selectionProtectionRuleModel = accessor.get(SelectionProtectionRuleModel);
        selectionProtectionRuleModel.setRule(unitId, subUnitId, ruleId, rule);
        return true;
    },
};
export const FactorySetSelectionProtection = (accessor: IAccessor, param: ISetSelectionProtectionParams) => {
    const { unitId, subUnitId, ruleId } = param;
    const selectionProtectionRuleModel = accessor.get(SelectionProtectionRuleModel);
    const oldRule = selectionProtectionRuleModel.getRule(unitId, subUnitId, ruleId);
    if (!oldRule) {
        return null;
    }
    const result: IMutationInfo<IAddSelectionProtectionParams> = {
        id: SetSelectionProtection.id,
        params: {
            ...param, rule: oldRule,
        },
    };
    return result;
};
