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

import type { IMutation, IMutationInfo } from '@univerjs/core';
import { CommandType } from '@univerjs/core';

import type { ISelectionProtectionRule } from '../../model/type';
import { SelectionProtectionRuleModel } from '../../model/selection-protection-rule.model';
import type { IDeleteSelectionProtectionParams } from './delete-selection-protection.mutation';
import { DeleteSelectionProtection } from './delete-selection-protection.mutation';

export interface IAddSelectionProtectionParams {
    rules: ISelectionProtectionRule[];
    unitId: string;
    subUnitId: string;
    name: string;
    description?: string;
};

export const FactoryAddSelectionProtection = (param: IAddSelectionProtectionParams) => {
    const deleteParams: IDeleteSelectionProtectionParams = { ...param, ruleIds: param.rules.map((rule) => rule.id) };
    return { id: DeleteSelectionProtection.id, params: deleteParams } as IMutationInfo<IDeleteSelectionProtectionParams>;
};

export const AddSelectionProtection: IMutation<IAddSelectionProtectionParams> = {
    id: 'sheet.mutation.add.selection-protection',
    type: CommandType.MUTATION,
    handler: (accessor, params) => {
        const { unitId, subUnitId, rules } = params;
        const selectionProtectionRuleModel = accessor.get(SelectionProtectionRuleModel);
        rules.forEach((rule) => {
            selectionProtectionRuleModel.addRule(unitId, subUnitId, rule);
        });
        return true;
    },
};
