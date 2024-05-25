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

import type { IPermissionPoint, IRangePermissionId } from '@univerjs/core';
import { PermissionStatus, PermissionType, RangeUnitPermissionType } from '@univerjs/core';

export class SelectionProtectionPermissionEditPoint implements IPermissionPoint {
    type = PermissionType.SHEET_RANGE;
    subType = RangeUnitPermissionType.Edit;
    status = PermissionStatus.INIT;
    value = true;
    id: IRangePermissionId;
    unitId: string;
    subUnitId: string;
    permissionId: string;
    constructor(unitId: string, subUnitId: string, permissionId: string) {
        this.unitId = unitId;
        this.subUnitId = subUnitId;
        this.permissionId = permissionId;
        this.id = `${PermissionType.SHEET_RANGE}.${RangeUnitPermissionType.Edit}.${permissionId}` as IRangePermissionId;
    }
}

export class SelectionProtectionPermissionViewPoint implements IPermissionPoint {
    type = PermissionType.SHEET_RANGE;
    subType = RangeUnitPermissionType.View;
    status = PermissionStatus.INIT;
    value = true;
    id: IRangePermissionId;
    unitId: string;
    subUnitId: string;
    permissionId: string;
    constructor(unitId: string, subUnitId: string, permissionId: string) {
        this.unitId = unitId;
        this.subUnitId = subUnitId;
        this.permissionId = permissionId;
        this.id = `${PermissionType.SHEET_RANGE}.${RangeUnitPermissionType.View}.${permissionId}` as IRangePermissionId;
    }
}

export type ISelectionPermissionPoint = SelectionProtectionPermissionViewPoint | SelectionProtectionPermissionEditPoint;

export const getAllRangePermissionPoint = () => [SelectionProtectionPermissionViewPoint, SelectionProtectionPermissionEditPoint];

export const getDefaultPermission = (unitId = 'unitId', subUnitId = 'subUnitId', permissionId = 'permissionId') => getAllRangePermissionPoint().reduce((r, F) => {
    const i = new F(unitId, subUnitId, permissionId);
    r[i.subType] = i.value;
    return r;
}, {} as Record<RangeUnitPermissionType, boolean>);
