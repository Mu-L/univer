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

import type { IRange, RangeUnitPermissionType, SubUnitPermissionType, UnitPermissionType } from '@univerjs/core';
import { PermissionType } from '@univerjs/core';

export function getIdByRange(range: IRange) {
    const { startRow, startColumn, endRow, endColumn } = range;
    return `${startRow}-${startColumn}-${endRow}-${endColumn}`;
}

export const getWorkbookPointId = (unitId: string, type: UnitPermissionType) => `${PermissionType.WORK_BOOK}.${type}_${unitId}`;
export const getWorksheetPointId = (unitId: string, subUnitId: string, type: SubUnitPermissionType) => `${PermissionType.WORK_SHEET}.${type}_${unitId}_${subUnitId}`;
export const getRangePointId = (permissionId: string, type: RangeUnitPermissionType) => `${PermissionType.SHEET_RANGE}.${type}.${permissionId}`;
