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

import type { IDrawingParam, Nullable } from '@univerjs/core';
import { ICommandService } from '@univerjs/core';
import { useDependency } from '@wendellhu/redi/react-bindings';
import React, { useRef, useState } from 'react';
import { ImageCommonPanel } from '@univerjs/image-ui';
import { ISheetDrawingService } from '@univerjs/sheets';


export const SheetImagePanel = () => {
    const commandService = useDependency(ICommandService);
    const sheetDrawingService = useDependency(ISheetDrawingService);
    const drawing = sheetDrawingService.getFocusDrawings()[0];

    if (drawing == null) {
        return;
    }

    const { unitId, subUnitId, drawingId, drawingType } = drawing;
    const props = {
        unitId, subUnitId, drawingId, drawingType,
    };

    return (
        <div className="container">
            <ImageCommonPanel {...props} />
        </div>
    );
};
