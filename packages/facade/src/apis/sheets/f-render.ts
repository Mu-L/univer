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

import { IRenderManagerService } from '@univerjs/engine-render';
import { SheetCanvasFloatDomManagerService } from '@univerjs/sheets-ui';
import { Inject } from '@wendellhu/redi';

export class FRender {
    constructor(
        @IRenderManagerService private _renderManagerService: IRenderManagerService,
        @Inject(SheetCanvasFloatDomManagerService) private _sheetCanvasFloatDomManagerService: SheetCanvasFloatDomManagerService
    ) {

    }

    addFloatDom() {}
}