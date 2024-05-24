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

import type { IRange, IScale, Worksheet } from '@univerjs/core';
import { Range } from '@univerjs/core';
import type { SpreadsheetSkeleton, UniverRenderingContext } from '@univerjs/engine-render';
import { SheetExtension } from '@univerjs/engine-render';
import type { ISelectionProtectionRenderCellData } from './type';
import { base64 } from './protect-background-img';

export const selectionProtectionKey = 'sheet-selection-protection';
const EXTENSION_Z_INDEX = 44;

export class SelectionProtectionRenderExtension extends SheetExtension {
    override uKey = selectionProtectionKey;

    override Z_INDEX = EXTENSION_Z_INDEX;
    private _pattern: CanvasPattern | null;

    private _img = new Image();

    public renderCache = new Set<string>();
    constructor() {
        super();
        this._img.src = base64;
    }

    override clearCache(): void {
        this.renderCache.clear();
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
        // 由于贴图渲染会将之前渲染的内容覆盖上去,这里无法做到增量渲染.
        ctx.save();
        if (!this._pattern) {
            this._pattern = ctx.createPattern(this._img, 'repeat');
        }
        this.renderCache.clear();
        Range.foreach(spreadsheetSkeleton.rowColumnSegment, (row, col) => {
            if (!worksheet.getColVisible(col) || !worksheet.getRowVisible(row)) {
                return;
            }
            const { selectionProtection = [] } = worksheet.getCell(row, col) as ISelectionProtectionRenderCellData || {};
            if (!this._pattern) {
                return;
            }
            ctx.fillStyle = this._pattern;
            selectionProtection.forEach((config) => {
                if (!config.ruleId) {
                    return;
                }
                if (this.renderCache.has(config.ruleId)) {
                    return;
                }
                this.renderCache.add(config.ruleId);
                config.ranges!.forEach((range) => {
                    const start = this.getCellIndex(range.startRow, range.startColumn, rowHeightAccumulation, columnWidthAccumulation, dataMergeCache);
                    const end = this.getCellIndex(range.endRow, range.endColumn, rowHeightAccumulation, columnWidthAccumulation, dataMergeCache);
                    ctx.fillRect(start.startX, start.startY, end.endX - start.startX, end.endY - start.startY);
                });
            });
        });
        ctx.restore();
    }

    isDirty(worksheet: Worksheet, ranges: IRange[]) {
        return ranges.some((range) => {
            let result = false;
            Range.foreach(range, (row: number, col: number) => {
                const { selectionProtection = [] } = worksheet.getCell(row, col) as ISelectionProtectionRenderCellData;
                for (const p of selectionProtection) {
                    if (p.ruleId && !this.renderCache.has(p.ruleId)) {
                        result = true;
                        return;
                    }
                }
            });
            return result;
        });
    }
}
