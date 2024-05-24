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

export { UniverSheetsSelectionProtectionPlugin } from './selection-protection-plugin';
export { SelectionProtectionRenderExtension, selectionProtectionKey } from './render/selection-protection.render';
export { getAllRangePermissionPoint, getDefaultPermission, SelectionProtectionPermissionEditPoint } from './service/selection-protection/permission-point';
export { AddSelectionProtection, DeleteSelectionProtection, SetSelectionProtection } from './commands';
export { AddRangeProtectionCommand } from './commands/commands/add-selection-protection.command';
export { DeleteRangeProtectionCommand } from './commands/commands/delete-selection-protection.command';
export { SetRangeProtectionCommand } from './commands/commands/set-selection-protection.command';
export { SelectionProtectionRenderModel } from './model/selection-protection-render.model';
export type { ISelectionProtectionRule, ICellPermission } from './model';
export { SelectionProtectionRuleModel } from './model';
export type { ISetSelectionProtectionParams } from './commands/mutation/set-selection-protection.mutation';
