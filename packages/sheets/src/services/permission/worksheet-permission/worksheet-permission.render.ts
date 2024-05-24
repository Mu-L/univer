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

import type { IRange, IScale } from '@univerjs/core';
import type { SpreadsheetSkeleton, UniverRenderingContext } from '@univerjs/engine-render';
import { SheetExtension } from '@univerjs/engine-render';
import type { IWorksheetProtectionRenderCellData } from './type';
import { base64 } from './protect-background-img';

export const worksheetProtectionKey = 'worksheet-protection';
const EXTENSION_Z_INDEX = 44;

export class WorksheetProtectionRenderExtension extends SheetExtension {
    override uKey = worksheetProtectionKey;

    override Z_INDEX = EXTENSION_Z_INDEX;
    private _pattern: CanvasPattern | null;

    private _img = new Image();

    constructor() {
        super();
        this._img.src = base64;
    }

    override draw(
        ctx: UniverRenderingContext,
        _parentScale: IScale,
        spreadsheetSkeleton: SpreadsheetSkeleton,
        _diffRanges?: IRange[]
    ) {
        const { rowHeightAccumulation, columnWidthAccumulation, worksheet, dataMergeCache } =
            spreadsheetSkeleton;
        if (!worksheet) {
            return false;
        }

        ctx.save();
        if (!this._pattern) {
            this._pattern = ctx.createPattern(this._img, 'repeat');
        }
        const { startRow, startColumn, endRow, endColumn } = spreadsheetSkeleton.rowColumnSegment;
        const start = this.getCellIndex(startRow, startColumn, rowHeightAccumulation, columnWidthAccumulation, dataMergeCache);
        const end = this.getCellIndex(endRow, endColumn, rowHeightAccumulation, columnWidthAccumulation, dataMergeCache);

        const { hasWorksheetRule = false } = worksheet.getCell(startRow, startColumn) as IWorksheetProtectionRenderCellData || {};
        if (!this._pattern) {
            return;
        }
        ctx.fillStyle = this._pattern;
        if (hasWorksheetRule) {
            ctx.fillRect(start.startX, start.startY, end.endX - start.startX, end.endY - start.startY);
        }

        ctx.restore();
    }
}
