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

import type { DependencyOverride } from '@univerjs/core';
import { ICommandService, mergeOverrideWithDependencies, Plugin, UniverInstanceType } from '@univerjs/core';
import type { Dependency } from '@wendellhu/redi';
import { Inject, Injector } from '@wendellhu/redi';
import { SelectionProtectionRenderModel } from './model/selection-protection-render.model';
import { SelectionProtectionRuleModel } from './model/selection-protection-rule.model';
import { SelectionProtectionService } from './service/selection-protection/selection-protection.service';
import { AddSelectionProtection } from './commands/mutation/add-selection-protection.mutation';
import { DeleteSelectionProtection } from './commands/mutation/delete-selection-protection.mutation';
import { SetSelectionProtection } from './commands/mutation/set-selection-protection.mutation';
import { PLUGIN_NAME } from './base/const';
import { AddRangeProtectionCommand } from './commands/commands/add-selection-protection.command';
import { SetRangeProtectionCommand } from './commands/commands/set-selection-protection.command';
import { DeleteRangeProtectionCommand } from './commands/commands/delete-selection-protection.command';
import { SelectionProtectionRefRangeService } from './service/selection-protection/selection-protection.ref-range';

export class UniverSheetsSelectionProtectionPlugin extends Plugin {
    static override type = UniverInstanceType.UNIVER_SHEET;
    static override pluginName = PLUGIN_NAME;

    static readonly dependencyList: Dependency[] = [[SelectionProtectionRenderModel], [SelectionProtectionRuleModel], [SelectionProtectionService], [SelectionProtectionRefRangeService]];
    static readonly mutationList = [AddSelectionProtection, DeleteSelectionProtection, SetSelectionProtection];
    static readonly commandList = [AddRangeProtectionCommand, DeleteRangeProtectionCommand, SetRangeProtectionCommand];

    constructor(
        private _config: {
            override?: DependencyOverride;
        },
        @Inject(Injector) override readonly _injector: Injector,
        @Inject(ICommandService) private _commandService: ICommandService

    ) {
        super();
        this._initCommand();
    }

    override onStarting(): void {
        const dependency = mergeOverrideWithDependencies(UniverSheetsSelectionProtectionPlugin.dependencyList, this._config?.override);
        dependency.forEach((d) => {
            this._injector.add(d);
        });
    }

    _initCommand() {
        UniverSheetsSelectionProtectionPlugin.mutationList.forEach((m) => {
            this._commandService.registerCommand(m);
        });
        UniverSheetsSelectionProtectionPlugin.commandList.forEach((c) => {
            this._commandService.registerCommand(c);
        });
    }
}
