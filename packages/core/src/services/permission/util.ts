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

import { UnitAction } from '@univerjs/protocol';
import { RangeUnitPermissionType, SubUnitPermissionType, UnitPermissionType } from './type';

type ISubType = SubUnitPermissionType | RangeUnitPermissionType | UnitPermissionType;
export function mapSubEnumToPermissionPoint(permissionEnum: UnitAction): ISubType {
    switch (permissionEnum) {
        case UnitAction.Copy:
            return SubUnitPermissionType.Copy;
        case UnitAction.SelectProtectedCells:
            return SubUnitPermissionType.SelectProtectedCells;
        case UnitAction.SelectUnProtectedCells:
            return SubUnitPermissionType.SelectUnProtectedCells;
        case UnitAction.SetCellStyle:
            return SubUnitPermissionType.SetCellStyle;
        case UnitAction.SetCellValue:
            return SubUnitPermissionType.SetCellValue;
        case UnitAction.View:
            return SubUnitPermissionType.View;
        case UnitAction.SetRowStyle:
            return SubUnitPermissionType.SetRowStyle;
        case UnitAction.SetColumnStyle:
            return SubUnitPermissionType.SetColumnStyle;
        case UnitAction.InsertRow:
            return SubUnitPermissionType.InsertRow;
        case UnitAction.InsertColumn:
            return SubUnitPermissionType.InsertColumn;
        case UnitAction.InsertHyperlink:
            return SubUnitPermissionType.InsertHyperlink;
        case UnitAction.DeleteRow:
            return SubUnitPermissionType.DeleteRow;
        case UnitAction.DeleteColumn:
            return SubUnitPermissionType.DeleteColumn;
        case UnitAction.Sort:
            return SubUnitPermissionType.Sort;
        case UnitAction.Filter:
            return SubUnitPermissionType.Filter;
        case UnitAction.PivotTable:
            return SubUnitPermissionType.PivotTable;
        case UnitAction.EditExtraObject:
            return SubUnitPermissionType.EditExtraObject;
        default:
            return SubUnitPermissionType.UnRecognized;
    }
}

export function mapPermissionPointToSubEnum(permissionPoint: ISubType): UnitAction {
    switch (permissionPoint) {
        case SubUnitPermissionType.Copy:
            return UnitAction.Copy;
        case SubUnitPermissionType.SelectProtectedCells:
            return UnitAction.SelectProtectedCells;
        case SubUnitPermissionType.SelectUnProtectedCells:
            return UnitAction.SelectUnProtectedCells;
        case SubUnitPermissionType.SetCellStyle:
            return UnitAction.SetCellStyle;
        case SubUnitPermissionType.SetCellValue:
            return UnitAction.SetCellValue;
        case SubUnitPermissionType.View:
            return UnitAction.View;
        case SubUnitPermissionType.SetRowStyle:
            return UnitAction.SetRowStyle;
        case SubUnitPermissionType.SetColumnStyle:
            return UnitAction.SetColumnStyle;
        case SubUnitPermissionType.InsertRow:
            return UnitAction.InsertRow;
        case SubUnitPermissionType.InsertColumn:
            return UnitAction.InsertColumn;
        case SubUnitPermissionType.InsertHyperlink:
            return UnitAction.InsertHyperlink;
        case SubUnitPermissionType.DeleteRow:
            return UnitAction.DeleteRow;
        case SubUnitPermissionType.DeleteColumn:
            return UnitAction.DeleteColumn;
        case SubUnitPermissionType.Sort:
            return UnitAction.Sort;
        case SubUnitPermissionType.Filter:
            return UnitAction.Filter;
        case SubUnitPermissionType.PivotTable:
            return UnitAction.PivotTable;
        case SubUnitPermissionType.EditExtraObject:
            return UnitAction.EditExtraObject;
        case RangeUnitPermissionType.Edit:
            return UnitAction.Edit;
        case UnitPermissionType.ManageCollaborator:
            return UnitAction.ManageCollaborator;
        default:
            return UnitAction.UNRECOGNIZED;
    }
}

