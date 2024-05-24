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

/* eslint-disable complexity */
/* eslint-disable max-lines-per-function */

import type { ICellData, ICellDataForSheetInterceptor, ICommandInfo, IObjectMatrixPrimitiveType, IPermissionTypes, IRange, ISheetDataValidationRule, Nullable, Workbook } from '@univerjs/core';
import { DisposableCollection, ICommandService, IUniverInstanceService, LifecycleStages, LocaleService, ObjectMatrix, OnLifecycle, RangeUnitPermissionType, Rectangle, RxDisposable, SubUnitPermissionType, Tools, UnitPermissionType, UniverInstanceType } from '@univerjs/core';
import type { EffectRefRangeParams, GetWorkbookPermissionFunc, GetWorksheetPermission, IInsertColCommandParams, IInsertColMutationParams, IInsertRowCommandParams, IMoveColsCommandParams, IMoveRangeCommandParams, IMoveRowsCommandParams, IMoveRowsMutationParams, IRemoveRowColCommandParams, ISetRangeValuesCommandParams, ISetSpecificColsVisibleCommandParams, ISetSpecificRowsVisibleCommandParams, ISetWorksheetActivateCommandParams, ISetWorksheetNameMutationParams } from '@univerjs/sheets';
import { ClearSelectionContentCommand, DeleteRangeMoveLeftCommand, DeleteRangeMoveUpCommand, DeltaColumnWidthCommand, DeltaRowHeightCommand, InsertColCommand, InsertColMutation, InsertRangeMoveDownCommand, InsertRangeMoveRightCommand, InsertRowCommand, InsertRowMutation, MoveColsCommand, MoveColsMutation, MoveRangeCommand, MoveRowsCommand, MoveRowsMutation, RefRangeService, RemoveColCommand, RemoveColMutation, RemoveRowCommand, RemoveRowMutation, SelectionManagerService, SetBackgroundColorCommand, SetColWidthCommand, SetRangeValuesCommand, SetRowHeightCommand, SetSelectedColsVisibleCommand, SetSelectedRowsVisibleCommand, SetSpecificColsVisibleCommand, SetSpecificRowsVisibleCommand, SetWorksheetActivateCommand, SetWorksheetNameCommand, SetWorksheetNameMutation, SetWorksheetOrderCommand, SetWorksheetRowIsAutoHeightCommand, SetWorksheetShowCommand, WorkbookPermissionService, WorksheetPermissionService, WorksheetProtectionRuleModel } from '@univerjs/sheets';
import { Inject } from '@wendellhu/redi';
import { IDialogService } from '@univerjs/ui';

import type { ISelectionProtectionRule, ISetSelectionProtectionParams } from '@univerjs/sheets-selection-protection';
import { AddRangeProtectionCommand, SelectionProtectionRuleModel, SetRangeProtectionCommand, SetSelectionProtection } from '@univerjs/sheets-selection-protection';

import type { IRenderContext, SpreadsheetSkeleton } from '@univerjs/engine-render';
import type { ISheetPasteParams } from '@univerjs/sheets-ui';
import { ApplyFormatPainterCommand, AutoFillCommand, HeaderMoveRenderController, HeaderResizeRenderController, IAutoFillService, ISelectionRenderService, ISheetClipboardService, SetCellEditVisibleOperation, SetRangeBoldCommand, SetRangeItalicCommand, SetRangeStrickThroughCommand, SetRangeUnderlineCommand, SheetCopyCommand, SheetCutCommand, SheetPasteColWidthCommand, SheetPasteCommand, SheetPasteShortKeyCommand, virtualizeDiscreteRanges } from '@univerjs/sheets-ui';
import { SheetsFilterService } from '@univerjs/sheets-filter';
import { OpenFilterPanelOperation } from '@univerjs/sheets-filter-ui';
import { SheetsFindReplaceController } from '@univerjs/sheets-find-replace';
import { InsertCommand } from '@univerjs/docs';
import type { IUpdateSheetDataValidationRangeCommandParams } from '@univerjs/sheets-data-validation';
import { AddSheetDataValidationCommand, DataValidationController, UpdateSheetDataValidationRangeCommand } from '@univerjs/sheets-data-validation';
import type { IAddCfCommandParams } from '@univerjs/sheets-conditional-formatting-ui';
import { AddCfCommand, ConditionalFormattingClearController } from '@univerjs/sheets-conditional-formatting-ui';
import type { IConditionalFormattingRuleConfig, IConditionFormattingRule } from '@univerjs/sheets-conditional-formatting';
import { serializeRange } from '@univerjs/engine-formula';
import { HeaderFreezeRenderController } from '@univerjs/sheets-ui/controllers/render-controllers/freeze.render-controller.js';
import { UNIVER_SHEET_PERMISSION_ALERT_DIALOG, UNIVER_SHEET_PERMISSION_ALERT_DIALOG_ID } from '../views/error-msg-dialog/interface';

type ICellPermission = Record<RangeUnitPermissionType, boolean> & { ruleId?: string; ranges?: IRange[] };
type ICheckPermissionCommandParams = IMoveRowsCommandParams | IMoveColsCommandParams | IMoveRangeCommandParams | ISetRangeValuesCommandParams | ISheetPasteParams | ISetSpecificRowsVisibleCommandParams | IUpdateSheetDataValidationRangeCommandParams | IAddCfCommandParams;
type IMoveRowsOrColsMutationParams = IMoveRowsMutationParams;

const mutationIdByRowCol = [InsertColMutation.id, InsertRowMutation.id, RemoveColMutation.id, RemoveRowMutation.id];
const mutationIdArrByMove = [MoveRowsMutation.id, MoveColsMutation.id];
const SmartToggleSheetsFilterCommandId = 'sheet.command.smart-toggle-filter';

export const SHEET_PERMISSION_PASTE_PLUGIN = 'SHEET_PERMISSION_PASTE_PLUGIN';

@OnLifecycle(LifecycleStages.Rendered, SheetPermissionInterceptorController)
export class SheetPermissionInterceptorController extends RxDisposable {
    disposableCollection = new DisposableCollection();

    constructor(
        private readonly _context: IRenderContext<Workbook>,
        @ICommandService private readonly _commandService: ICommandService,
        @IUniverInstanceService private readonly _univerInstanceService: IUniverInstanceService,
        @Inject(WorkbookPermissionService) private readonly _workbookPermissionService: WorkbookPermissionService,
        @Inject(WorksheetPermissionService) private readonly _worksheetPermissionService: WorksheetPermissionService,
        @Inject(SelectionManagerService) private readonly _selectionManagerService: SelectionManagerService,
        @IDialogService private readonly _dialogService: IDialogService,
        @Inject(SelectionProtectionRuleModel) private _selectionProtectionRuleModel: SelectionProtectionRuleModel,
        @Inject(WorksheetProtectionRuleModel) private _worksheetProtectionRuleModel: WorksheetProtectionRuleModel,
        @Inject(RefRangeService) private readonly _refRangeService: RefRangeService,
        @Inject(HeaderMoveRenderController) private _headerMoveRenderController: HeaderMoveRenderController,
        @Inject(HeaderResizeRenderController) private _headerResizeRenderController: HeaderResizeRenderController,
        @ISelectionRenderService private _selectionRenderService: ISelectionRenderService,
        @IAutoFillService private _autoFillService: IAutoFillService,
        @Inject(SheetsFilterService) private _sheetsFilterService: SheetsFilterService,
        @Inject(SheetsFindReplaceController) private _sheetsFindReplaceController: SheetsFindReplaceController,
        @Inject(LocaleService) private readonly _localService: LocaleService,
        @Inject(DataValidationController) private readonly _dataValidationController: DataValidationController,
        @Inject(ConditionalFormattingClearController) private readonly _conditionalFormattingClearController: ConditionalFormattingClearController,
        @Inject(ISheetClipboardService) private _sheetClipboardService: ISheetClipboardService,
        @Inject(HeaderFreezeRenderController) private _headerFreezeRenderController: HeaderFreezeRenderController
    ) {
        super();
        this._initialize();
        this._onRefRangeChange();
        this._correctPermissionRange();
        this._initHeaderMovePermissionInterceptor();
        this._initHeaderResizePermissionInterceptor();
        this._initRangeFillPermissionInterceptor();
        this._initRangeMovePermissionInterceptor();
        this._initSheetFindPermissionInterceptor();
        this._initDataValidationPermissionInterceptor();
        this._initConditionalFormattingPermissionInterceptor();
        this._initFreezePermissionInterceptor();
        this._initClipboardHook();
    }

    private _haveNotPermissionHandle(errorMsg: string) {
        const dialogProps = {
            id: UNIVER_SHEET_PERMISSION_ALERT_DIALOG_ID,
            title: { title: '' },
            children: {
                label: UNIVER_SHEET_PERMISSION_ALERT_DIALOG,
                errorMsg,
            },
            width: 320,
            destroyOnClose: true,
            onClose: () => this._dialogService.close(UNIVER_SHEET_PERMISSION_ALERT_DIALOG_ID),
            className: 'sheet-permission-user-dialog',
        };
        this._dialogService.open(dialogProps);
        throw new Error('have not permission');
    }

    private _getPermissionCheck(id: string, params: ICheckPermissionCommandParams) {
        let permission = true;
        let errorMsg = '';
        switch (id) {
            case InsertCommand.id:
            case SetCellEditVisibleOperation.id:
                permission = this._permissionCheckWithoutRange({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetCellValue, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.editErr');
                break;
            case SetRangeValuesCommand.id:
                permission = this._permissionCheckBySetRangeValue({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetCellValue, SubUnitPermissionType.Edit],
                }, params as ISetRangeValuesCommandParams);
                break;
            case ClearSelectionContentCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetCellValue, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.editErr');
                break;
            case SheetPasteColWidthCommand.id:
                permission = this._permissionCheckWithoutRange({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetColumnStyle],
                });
                errorMsg = this._localService.t('permission.dialog.pasteErr');
                break;
            case SheetPasteCommand.id:
            case SheetPasteShortKeyCommand.id:
                permission = this._permissionCheckByPaste(params as ISheetPasteParams);
                errorMsg = this._localService.t('permission.dialog.pasteErr');
                break;
            case ApplyFormatPainterCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetCellValue, SubUnitPermissionType.SetCellStyle],
                });
                errorMsg = this._localService.t('permission.dialog.commonErr');
                break;
            case SetBackgroundColorCommand.id:
            case SetRangeBoldCommand.id:
            case SetRangeItalicCommand.id:
            case SetRangeUnderlineCommand.id:
            case SetRangeStrickThroughCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetCellStyle, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.setStyleErr');
                break;
            case SheetCopyCommand.id:
            case SheetCutCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Copy],
                    rangeTypes: [RangeUnitPermissionType.View],
                    worksheetTypes: [SubUnitPermissionType.Copy],
                });
                errorMsg = this._localService.t('permission.dialog.copyErr');
                break;
            case DeltaColumnWidthCommand.id:
            case SetColWidthCommand.id:
                permission = this._permissionCheckWithoutRange({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetRowStyle, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;

            case DeltaRowHeightCommand.id:
            case SetRowHeightCommand.id:
            case SetWorksheetRowIsAutoHeightCommand.id:
                permission = this._permissionCheckWithoutRange({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.SetRowStyle, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;
            case MoveColsCommand.id:
            case MoveRowsCommand.id:
                permission = this._permissionCheckByMoveCommand(params as IMoveColsCommandParams);
                errorMsg = this._localService.t('permission.dialog.moveRowColErr');
                break;

            case MoveRangeCommand.id:
                permission = this._permissionCheckByMoveRangeCommand(params as IMoveRangeCommandParams);
                errorMsg = this._localService.t('permission.dialog.moveRangeErr');
                break;

            case AutoFillCommand.id:
                permission = this._permissionCheckByAutoFillCommand(this._autoFillService.autoFillLocation?.target);
                errorMsg = this._localService.t('permission.dialog.autoFillErr');
                break;

            case SmartToggleSheetsFilterCommandId:
                permission = this._permissionCheckWithoutRange({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.View],
                    worksheetTypes: [SubUnitPermissionType.Filter, SubUnitPermissionType.Edit],
                });
                errorMsg = this._localService.t('permission.dialog.filterErr');
                break;
            case SetWorksheetOrderCommand.id:
            case SetWorksheetNameCommand.id:
            case SetWorksheetShowCommand.id:
                permission = this._permissionCheckByWorksheetCommand();
                errorMsg = this._localService.t('permission.dialog.operatorSheetErr');
                if (permission === false) {
                    this._worksheetProtectionRuleModel.resetOrder();
                }
                break;
            case OpenFilterPanelOperation.id:
                permission = this._permissionCheckWithFilter();
                errorMsg = this._localService.t('permission.dialog.filterErr');
                break;

            case SetSpecificColsVisibleCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetColumnStyle],
                }, (params as ISetSpecificColsVisibleCommandParams).ranges);
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;
            case SetSpecificRowsVisibleCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetRowStyle],
                }, (params as ISetSpecificRowsVisibleCommandParams).ranges);
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;
            case SetSelectedColsVisibleCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetColumnStyle],
                });
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;
            case SetSelectedRowsVisibleCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetRowStyle],
                });
                errorMsg = this._localService.t('permission.dialog.setRowColStyleErr');
                break;

            case AddSheetDataValidationCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetCellStyle],
                });
                errorMsg = this._localService.t('permission.dialog.setStyleErr');
                break;
            case UpdateSheetDataValidationRangeCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetCellStyle],
                }, (params as IUpdateSheetDataValidationRangeCommandParams).ranges);
                errorMsg = this._localService.t('permission.dialog.setStyleErr');
                break;

            case AddCfCommand.id:
                permission = this._permissionCheckWithRanges({
                    workbookTypes: [UnitPermissionType.Edit],
                    rangeTypes: [RangeUnitPermissionType.Edit],
                    worksheetTypes: [SubUnitPermissionType.Edit, SubUnitPermissionType.SetCellStyle],
                }, (params as IAddCfCommandParams).rule.ranges);
                errorMsg = this._localService.t('permission.dialog.setStyleErr');
                break;

            case InsertRangeMoveRightCommand.id:
                permission = this._permissionCheckWithInsertRangeMove('right');
                errorMsg = this._localService.t('permission.dialog.insertOrDeleteMoveRangeErr');
                break;
            case InsertRangeMoveDownCommand.id:
                permission = this._permissionCheckWithInsertRangeMove('bottom');
                errorMsg = this._localService.t('permission.dialog.insertOrDeleteMoveRangeErr');
                break;
            case DeleteRangeMoveLeftCommand.id:
                permission = this._permissionCheckWithInsertRangeMove('left');
                errorMsg = this._localService.t('permission.dialog.insertOrDeleteMoveRangeErr');
                break;
            case DeleteRangeMoveUpCommand.id:
                permission = this._permissionCheckWithInsertRangeMove('top');
                errorMsg = this._localService.t('permission.dialog.insertOrDeleteMoveRangeErr');
                break;
            default:
                break;
        }

        if (!permission) {
            this._haveNotPermissionHandle(errorMsg);
        }
    };

    private _initialize(): void {
        this._commandExecutedListener();
    }

    private _commandExecutedListener() {
        this.disposeWithMe(
            this._commandService.beforeCommandExecuted((command: ICommandInfo) => {
                this._getPermissionCheck(command.id, command?.params as ICheckPermissionCommandParams);
            })
        );
        this.disposeWithMe(
            this._commandService.onCommandExecuted((command: ICommandInfo) => {
                if (command.id === SetWorksheetNameMutation.id) {
                    const params = command.params as ISetWorksheetNameMutationParams;
                    const { unitId = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)?.getUnitId(), subUnitId } = params;
                    if (!unitId || !subUnitId) {
                        return;
                    }
                    const worksheetRule = this._worksheetProtectionRuleModel.getRule(unitId, subUnitId);
                    const selectionRuleList = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId);
                    if (worksheetRule) {
                        worksheetRule.name = params.name;
                        this._worksheetProtectionRuleModel.ruleRefresh(worksheetRule.permissionId);
                    }
                    if (selectionRuleList.length) {
                        selectionRuleList.forEach((rule) => {
                            const ranges = rule.ranges;
                            const rangeStr = ranges?.length
                                ? ranges.map((range) => {
                                    const v = serializeRange(range);
                                    return v === 'NaN' ? '' : v;
                                }).filter((r) => !!r).join(',')
                                : '';
                            rule.name = `${params.name}(${rangeStr})`;
                        });
                        this._selectionProtectionRuleModel.ruleRefresh(subUnitId);
                    }
                }
            })
        );
    }

    private _permissionCheckWithInsertRangeMove(direction: 'top' | 'bottom' | 'left' | 'right') {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const unitId = workbook.getUnitId();
        const worksheet = workbook?.getActiveSheet();
        const subUnitId = worksheet.getSheetId();
        const selectionRange = this._selectionManagerService.getLast()?.range;
        if (!selectionRange) {
            return false;
        };
        if (direction === 'top') {
            selectionRange.startRow = 0;
        } else if (direction === 'left') {
            selectionRange.startColumn = 0;
        } else if (direction === 'right') {
            selectionRange.endColumn = worksheet.getColumnCount() - 1;
        } else if (direction === 'bottom') {
            selectionRange.endRow = worksheet.getRowCount() - 1;
        }

        const selectionRuleRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).map((rule) => rule.ranges).flat();
        const hasLap = selectionRuleRanges.some((range) => {
            return Rectangle.getIntersects(selectionRange, range);
        });
        if (hasLap) {
            return false;
        }
        return true;
    }

    private _permissionCheckWithFilter() {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const unitId = workbook.getUnitId();
        const worksheet = workbook?.getActiveSheet();
        const subUnitId = worksheet.getSheetId();
        const filterRange = this._sheetsFilterService.getFilterModel(unitId, subUnitId)?.getRange();
        if (filterRange) {
            return this._permissionCheckWithRanges({
                workbookTypes: [UnitPermissionType.Edit],
                rangeTypes: [RangeUnitPermissionType.View],
                worksheetTypes: [SubUnitPermissionType.Filter, SubUnitPermissionType.Edit],
            }, [filterRange]);
        }
        return true;
    }

    private _permissionCheckByWorksheetCommand() {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const unitId = workbook.getUnitId();
        const worksheet = workbook?.getActiveSheet();
        const subUnitId = worksheet.getSheetId();
        const worksheetRule = this._worksheetProtectionRuleModel.getRule(unitId, subUnitId);
        if (worksheetRule) {
            return this._workbookPermissionService.getManageCollaboratorPermission(unitId);
        } else {
            return this._workbookPermissionService.getEditPermission(unitId);
        }
    }

    private _permissionCheckWithoutRange(permissionTypes: IPermissionTypes) {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const worksheet = workbook?.getActiveSheet();
        const selection = this._selectionManagerService.getLast();
        const row = selection?.primary?.actualRow ?? 0;
        const col = selection?.primary?.actualColumn ?? 0;
        const { workbookTypes = [UnitPermissionType.Edit], worksheetTypes, rangeTypes } = permissionTypes;
        if (workbookTypes) {
            const workbookDisable = workbookTypes.some((type) => {
                const workbookPermissionCheckFnName = `get${type}Permission` as keyof WorkbookPermissionService;
                const workbookPermissionCheckFn = this._workbookPermissionService[workbookPermissionCheckFnName] as GetWorkbookPermissionFunc;
                const workbookPermission = workbookPermissionCheckFn(workbook.getUnitId());
                return workbookPermission === false;
            });
            if (workbookDisable === true) {
                return false;
            }
        }
        if (worksheetTypes) {
            const worksheetDisable = worksheetTypes.some((type) => {
                const worksheetPermissionCheckFnName = `get${type}Permission` as keyof WorksheetPermissionService;
                const worksheetPermissionCheckFn = this._worksheetPermissionService[worksheetPermissionCheckFnName] as GetWorksheetPermission;
                const worksheetPermission = worksheetPermissionCheckFn({
                    unitId: workbook.getUnitId(),
                    subUnitId: worksheet.getSheetId(),
                });
                return worksheetPermission === false;
            });
            if (worksheetDisable === true) {
                return false;
            }
        }
        if (rangeTypes) {
            const rangeDisable = rangeTypes.some((rangeType) => {
                const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                return permission?.[rangeType] === false;
            });
            if (rangeDisable === true) {
                return false;
            }
        }
        return true;
    }

    private _permissionCheckWithRanges(permissionTypes: IPermissionTypes, selectionRanges?: IRange[], unitId?: string, subUnitId?: string) {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const worksheet = workbook?.getActiveSheet();
        const ranges = selectionRanges ?? this._selectionManagerService.getSelections()?.map((selection) => {
            return selection.range;
        });

        if (!ranges) {
            return false;
        }

        const { workbookTypes = [UnitPermissionType.Edit], worksheetTypes, rangeTypes } = permissionTypes;
        if (workbookTypes) {
            const workbookDisable = workbookTypes.some((type) => {
                const workbookPermissionCheckFnName = `get${type}Permission` as keyof WorkbookPermissionService;
                const workbookPermissionCheckFn = this._workbookPermissionService[workbookPermissionCheckFnName] as GetWorkbookPermissionFunc;
                const workbookPermission = workbookPermissionCheckFn(workbook.getUnitId());
                return workbookPermission === false;
            });
            if (workbookDisable === true) {
                return false;
            }
        }
        if (worksheetTypes) {
            const worksheetDisable = worksheetTypes.some((type) => {
                const worksheetPermissionCheckFnName = `get${type}Permission` as keyof WorksheetPermissionService;
                const worksheetPermissionCheckFn = this._worksheetPermissionService[worksheetPermissionCheckFnName] as GetWorksheetPermission;
                const worksheetPermission = worksheetPermissionCheckFn({
                    unitId: workbook.getUnitId(),
                    subUnitId: subUnitId ?? worksheet.getSheetId(),
                });
                return worksheetPermission === false;
            });
            if (worksheetDisable === true) {
                return false;
            }
        }
        if (rangeTypes) {
            const hasPermission = ranges?.every((range) => {
                return rangeTypes.every((rangeType) => {
                    for (let row = range.startRow; row <= range.endRow; row++) {
                        for (let col = range.startColumn; col <= range.endColumn; col++) {
                            const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                            if (permission?.[rangeType] === false) {
                                return false;
                            }
                        }
                    }
                    return true;
                });
            });
            return hasPermission;
        }

        return true;
    }

    private _permissionCheckByPaste(params: ISheetPasteParams) {
        if (params.value === 'special-paste-value' || params.value === 'special-paste-formula') {
            return this._permissionCheckWithRanges({
                workbookTypes: [UnitPermissionType.Edit],
                rangeTypes: [RangeUnitPermissionType.Edit],
                worksheetTypes: [SubUnitPermissionType.SetCellValue, SubUnitPermissionType.Edit],
            });
        } else if (params.value === 'special-paste-format') {
            return this._permissionCheckWithRanges({
                workbookTypes: [UnitPermissionType.Edit],
                rangeTypes: [RangeUnitPermissionType.Edit],
                worksheetTypes: [SubUnitPermissionType.SetCellStyle, SubUnitPermissionType.Edit],
            });
        } else {
            return this._permissionCheckWithRanges({
                workbookTypes: [UnitPermissionType.Edit],
                rangeTypes: [RangeUnitPermissionType.Edit],
                worksheetTypes: [SubUnitPermissionType.SetCellValue, SubUnitPermissionType.SetCellStyle, SubUnitPermissionType.Edit],
            });
        }
    }

    private _permissionCheckByMoveCommand(params: IMoveRowsCommandParams | IMoveColsCommandParams) {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const worksheet = workbook?.getActiveSheet();
        const unitId = workbook.getUnitId();
        const subUnitId = worksheet.getSheetId();
        const toRange = params.toRange;
        if (toRange.endRow === worksheet.getRowCount() - 1) {
            toRange.endColumn = toRange.startColumn;
        } else {
            toRange.endRow = toRange.startRow;
        }
        const permissionLapRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
            return [...p, ...c.ranges];
        }, [] as IRange[]).filter((range) => {
            return Rectangle.intersects(range, toRange);
        });

        if (permissionLapRanges.length > 0) {
            return false;
        }
        permissionLapRanges.forEach((range) => {
            for (let row = range.startRow; row <= range.endRow; row++) {
                for (let col = range.startColumn; col <= range.endColumn; col++) {
                    const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                    if (permission?.Edit === false) {
                        return false;
                    }
                }
            }
        });
        return true;
    }

    private _permissionCheckByAutoFillCommand(params?: { rows: number[]; cols: number[] }) {
        if (!params) {
            return false;
        }
        const { rows, cols } = params;
        const startRow = rows[0];
        const endRow = rows[rows.length - 1];
        const startCol = cols[0];
        const endCol = cols[cols.length - 1];

        const targetRange = { startRow, endRow, startColumn: startCol, endColumn: endCol };

        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const worksheet = workbook?.getActiveSheet();
        const unitId = workbook.getUnitId();
        const subUnitId = worksheet.getSheetId();

        const permissionLapRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
            return [...p, ...c.ranges];
        }, [] as IRange[]).filter((range) => {
            return Rectangle.intersects(range, targetRange);
        });

        const hasNotPermission = permissionLapRanges.some((range) => {
            for (let row = range.startRow; row <= range.endRow; row++) {
                for (let col = range.startColumn; col <= range.endColumn; col++) {
                    const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                    if (permission?.Edit === false) {
                        return true;
                    }
                }
            }
            return false;
        });

        return !hasNotPermission;
    }

    private _permissionCheckByMoveRangeCommand(params: IMoveRangeCommandParams) {
        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        const worksheet = workbook?.getActiveSheet();
        const unitId = workbook.getUnitId();
        const subUnitId = worksheet.getSheetId();
        const toRange = params.toRange;
        const permissionLapRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
            return [...p, ...c.ranges];
        }, [] as IRange[]).filter((range) => {
            return Rectangle.intersects(range, toRange);
        });

        if (permissionLapRanges.length > 0) {
            return false;
        }
        permissionLapRanges.forEach((range) => {
            for (let row = range.startRow; row <= range.endRow; row++) {
                for (let col = range.startColumn; col <= range.endColumn; col++) {
                    const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                    if (permission?.Edit === false) {
                        return false;
                    }
                }
            }
        });
        return true;
    }

    private _permissionCheckBySetRangeValue(permissionTypes: IPermissionTypes, setRangeValueParams: ISetRangeValuesCommandParams) {
        let ranges: IRange[] = [];
        if (setRangeValueParams.range) {
            ranges = [setRangeValueParams.range];
        } else {
            const matrix = new ObjectMatrix(setRangeValueParams.value as IObjectMatrixPrimitiveType<ICellData>);
            const range = matrix.getDataRange();
            ranges = [range];
        }
        const { unitId, subUnitId } = setRangeValueParams;
        return this._permissionCheckWithRanges(permissionTypes, ranges, unitId, subUnitId);
    }

    refRangeHandle(config: EffectRefRangeParams, unitId: string, subUnitId: string) {
        switch (config.id) {
            case MoveRowsCommand.id:
                return this._getRefRangeMutationsByMoveRows(config.params as IMoveRowsCommandParams, unitId, subUnitId);
            case MoveColsCommand.id:
                return this._getRefRangeMutationsByMoveCols(config.params as IMoveColsCommandParams, unitId, subUnitId);
            case InsertRowCommand.id:
                return this._getRefRangeMutationsByInsertRows(config.params as IInsertRowCommandParams, unitId, subUnitId);
            case InsertColCommand.id:
                return this._getRefRangeMutationsByInsertCols(config.params as IInsertColCommandParams, unitId, subUnitId);
            case RemoveColCommand.id:
                return this._getRefRangeMutationsByDeleteCols(config.params as IRemoveRowColCommandParams, unitId, subUnitId);
            case RemoveRowCommand.id:
                return this._getRefRangeMutationsByDeleteRows(config.params as IRemoveRowColCommandParams, unitId, subUnitId);
            default:
                break;
        }
        return { redos: [], undos: [] };
    }

    private _getRefRangeMutationsByDeleteCols(params: IRemoveRowColCommandParams, unitId: string, subUnitId: string) {
        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return Rectangle.intersects(range, params.range);
            });
        });

        const removeRange = params.range;
        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                const rangesByRemove = cloneRule.ranges.reduce((p, c) => {
                    if (Rectangle.intersects(c, removeRange)) {
                        const cloneRange = Tools.deepClone(c);
                        const { startColumn, endColumn } = removeRange;
                        if (startColumn <= cloneRange.startColumn && endColumn >= cloneRange.endColumn) {
                            return p;
                        } else if (startColumn >= cloneRange.startColumn && endColumn <= cloneRange.endColumn) {
                            cloneRange.endColumn -= endColumn - startColumn + 1;
                        } else if (startColumn < cloneRange.startColumn) {
                            cloneRange.startColumn = startColumn;
                            cloneRange.endColumn -= endColumn - startColumn + 1;
                        } else if (endColumn > cloneRange.endColumn) {
                            cloneRange.endColumn = startColumn - 1;
                        }
                        if (this._checkIsRightRange(cloneRange)) {
                            p.push(cloneRange);
                        }
                    }
                    return p;
                }, [] as IRange[]);
                cloneRule.ranges = rangesByRemove;
                redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
            });

            return { redos: redoMutations, undos: undoMutations };
        }
        return { undos: [], redos: [] };
    }

    private _getRefRangeMutationsByDeleteRows(params: IRemoveRowColCommandParams, unitId: string, subUnitId: string) {
        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return Rectangle.intersects(range, params.range);
            });
        });

        const removeRange = params.range;
        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                const rangesByRemove = cloneRule.ranges.reduce((p, c) => {
                    if (Rectangle.intersects(c, removeRange)) {
                        const cloneRange = Tools.deepClone(c);
                        const { startRow, endRow } = removeRange;
                        if (startRow <= cloneRange.startRow && endRow >= cloneRange.endRow) {
                            return p;
                        } else if (startRow >= cloneRange.startRow && endRow <= cloneRange.endRow) {
                            cloneRange.endRow -= endRow - startRow + 1;
                        } else if (startRow < cloneRange.startRow) {
                            cloneRange.startRow = startRow;
                            cloneRange.endRow -= endRow - startRow + 1;
                        } else if (endRow > cloneRange.endRow) {
                            cloneRange.endRow = startRow - 1;
                        }
                        if (this._checkIsRightRange(cloneRange)) {
                            p.push(cloneRange);
                        }
                    }
                    return p;
                }, [] as IRange[]);
                cloneRule.ranges = rangesByRemove;
                redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
            });

            return { redos: redoMutations, undos: undoMutations };
        }
        return { undos: [], redos: [] };
    }

    private _getRefRangeMutationsByInsertCols(params: IInsertColCommandParams, unitId: string, subUnitId: string) {
        const insertStart = params.range.startColumn;
        const insertLength = params.range.endColumn - params.range.startColumn + 1;

        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return insertStart > range.startColumn && insertStart <= range.endColumn;
            });
        });

        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                let hasLap = false;
                cloneRule.ranges.forEach((range) => {
                    if (insertStart > range.startColumn && insertStart <= range.endColumn) {
                        range.endColumn += insertLength;
                        hasLap = true;
                    }
                });
                if (hasLap) {
                    redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                    undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
                }
            });

            return { redos: redoMutations, undos: undoMutations };
        }
        return { undos: [], redos: [] };
    }

    private _getRefRangeMutationsByInsertRows(params: IInsertRowCommandParams, unitId: string, subUnitId: string) {
        const insertStart = params.range.startRow;
        const insertLength = params.range.endRow - params.range.startRow + 1;

        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return insertStart > range.startRow && insertStart <= range.endRow;
            });
        });

        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                let hasLap = false;
                cloneRule.ranges.forEach((range) => {
                    if (insertStart > range.startRow && insertStart <= range.endRow) {
                        range.endRow += insertLength;
                        hasLap = true;
                    }
                });
                if (hasLap) {
                    redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                    undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
                }
            });

            return { redos: redoMutations, undos: undoMutations };
        }
        return { undos: [], redos: [] };
    }

    private _getRefRangeMutationsByMoveRows(params: IMoveRowsCommandParams, unitId: string, subUnitId: string) {
        const toRange = params.toRange;
        const moveToStartRow = toRange.startRow;
        const moveLength = toRange.endRow - toRange.startRow + 1;

        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return moveToStartRow > range.startRow && moveToStartRow <= range.endRow;
            });
        });

        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                const fromRange = params.fromRange;
                const moveFromStartRow = fromRange.startRow;
                let hasLap = false;
                cloneRule.ranges.forEach((range) => {
                    if (moveToStartRow > range.startRow && moveToStartRow <= range.endRow) {
                        if (moveFromStartRow < range.startRow) {
                            range.startRow = range.startRow - moveLength;
                            range.endRow = range.endRow - moveLength;
                        }
                        range.endRow += moveLength;
                        hasLap = true;
                    }
                });
                if (hasLap) {
                    redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                    undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
                }
            });

            return { redos: redoMutations, undos: undoMutations };
        }

        return { undos: [], redos: [] };
    }

    private _getRefRangeMutationsByMoveCols(params: IMoveColsCommandParams, unitId: string, subUnitId: string) {
        const toRange = params.toRange;
        const moveToStartCol = toRange.startColumn;
        const moveLength = toRange.endColumn - toRange.startColumn + 1;

        const permissionRangeLapRules = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).filter((rule) => {
            return rule.ranges.some((range) => {
                return moveToStartCol > range.startColumn && moveToStartCol <= range.endColumn;
            });
        });

        if (permissionRangeLapRules.length) {
            const redoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            const undoMutations: { id: string; params: ISetSelectionProtectionParams }[] = [];
            permissionRangeLapRules.forEach((rule) => {
                const cloneRule = Tools.deepClone(rule);
                const fromRange = params.fromRange;
                const moveFromStartCol = fromRange.startColumn;
                let hasLap = false;
                cloneRule.ranges.forEach((range) => {
                    if (moveToStartCol > range.startColumn && moveToStartCol <= range.endColumn) {
                        if (moveFromStartCol < range.startColumn) {
                            range.startColumn = range.startColumn - moveLength;
                            range.endColumn = range.endColumn - moveLength;
                        }
                        range.endColumn += moveLength;
                        hasLap = true;
                    }
                });
                if (hasLap) {
                    redoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule: cloneRule, ruleId: rule.id } });
                    undoMutations.push({ id: SetSelectionProtection.id, params: { unitId, subUnitId, rule, ruleId: rule.id } });
                }
            });

            return { redos: redoMutations, undos: undoMutations };
        }

        return { undos: [], redos: [] };
    }

    private _onRefRangeChange() {
        const registerRefRange = (unitId: string, subUnitId: string) => {
            const workbook = this._univerInstanceService.getUniverSheetInstance(unitId);
            if (!workbook) {
                return;
            }
            const workSheet = workbook?.getSheetBySheetId(subUnitId);
            if (!workSheet) {
                return;
            }

            this.disposableCollection.dispose();

            const handler = (config: EffectRefRangeParams) => {
                return this.refRangeHandle(config, unitId, subUnitId);
            };

            const permissionRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                return [...p, ...c.ranges];
            }, [] as IRange[]);

            permissionRanges.forEach((range) => {
                this.disposableCollection.add(this._refRangeService.registerRefRange(range, handler, unitId, subUnitId));
            });
        };
        this.disposeWithMe(
            this._commandService.onCommandExecuted((commandInfo) => {
                if (commandInfo.id === SetWorksheetActivateCommand.id) {
                    const params = commandInfo.params as ISetWorksheetActivateCommandParams;
                    const sheetId = params.subUnitId;
                    const unitId = params.unitId;
                    if (!sheetId || !unitId) {
                        return;
                    }
                    registerRefRange(unitId, sheetId);
                }
                if (commandInfo.id === AddRangeProtectionCommand.id || commandInfo.id === SetRangeProtectionCommand.id) {
                    const params = commandInfo.params as {
                        permissionId: string;
                        rule: ISelectionProtectionRule;
                    };
                    const subUnitId = params.rule.subUnitId;
                    const unitId = params.rule.unitId;
                    if (!subUnitId || !unitId) {
                        return;
                    }
                    registerRefRange(unitId, subUnitId);
                }
            })
        );

        const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
        if (workbook) {
            const sheet = workbook.getActiveSheet();
            registerRefRange(workbook.getUnitId(), sheet.getSheetId());
        }
    }

    private _correctPermissionRange() {
        this.disposeWithMe(this._commandService.onCommandExecuted((command: ICommandInfo) => {
            if (mutationIdArrByMove.includes(command.id)) {
                if (!command.params) return;
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                if (!workbook) return;
                const worksheet = workbook.getSheetBySheetId((command.params as IMoveRowsMutationParams).subUnitId);
                if (!worksheet) return;
                const { sourceRange, targetRange } = command.params as IMoveRowsOrColsMutationParams;
                const isRowMove = sourceRange.startColumn === targetRange.startColumn && sourceRange.endColumn === targetRange.endColumn;
                const moveLength = isRowMove
                    ? sourceRange.endRow - sourceRange.startRow + 1
                    : sourceRange.endColumn - sourceRange.startColumn + 1;
                const sourceStart = isRowMove ? sourceRange.startRow : sourceRange.startColumn;
                const targetStart = isRowMove ? targetRange.startRow : targetRange.startColumn;
                const permissionListRule = this._selectionProtectionRuleModel.getSubunitRuleList(workbook.getUnitId(), worksheet.getSheetId());

                permissionListRule.forEach((rule) => {
                    const ranges = rule.ranges;
                    ranges.forEach((range) => {
                        let { startRow, endRow, startColumn, endColumn } = range;

                        if (!Rectangle.intersects(range, sourceRange)) {
                            if (isRowMove) {
                                if (sourceStart < startRow && targetStart > endRow) {
                                    startRow -= moveLength;
                                    endRow -= moveLength;
                                } else if (sourceStart > endRow && targetStart <= startRow) {
                                    startRow += moveLength;
                                    endRow += moveLength;
                                }
                            } else {
                                if (sourceStart < startColumn && targetStart > endColumn) {
                                    startColumn -= moveLength;
                                    endColumn -= moveLength;
                                } else if (sourceStart > endColumn && targetStart <= startColumn) {
                                    startColumn += moveLength;
                                    endColumn += moveLength;
                                }
                            }
                        }

                        if (this._checkIsRightRange({ startRow, endRow, startColumn, endColumn })) {
                            range.startColumn = startColumn;
                            range.endColumn = endColumn;
                            range.startRow = startRow;
                            range.endRow = endRow;
                        }
                    });
                });

                this.disposableCollection.dispose();
                const { unitId, subUnitId } = command.params as IMoveRowsMutationParams;
                const handler = (config: EffectRefRangeParams) => {
                    return this.refRangeHandle(config, unitId, subUnitId);
                };

                const permissionRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                    return [...p, ...c.ranges];
                }, [] as IRange[]);

                permissionRanges.forEach((range) => {
                    this.disposableCollection.add(this._refRangeService.registerRefRange(range, handler, unitId, subUnitId));
                });
            }

            // 2. InsertRowsOrCols / RemoveRowsOrCols Mutations
            if (mutationIdByRowCol.includes(command.id)) {
                const workbook = this._univerInstanceService.getUniverSheetInstance((command.params as IInsertColMutationParams).unitId);
                if (!workbook) return;
                const worksheet = workbook.getSheetBySheetId((command.params as IInsertColMutationParams).subUnitId);
                if (!worksheet) return;

                const params = command.params as IInsertRowCommandParams;
                if (!params) return;
                const { range } = params;

                const isRowOperation = command.id.includes('row');
                const isAddOperation = command.id.includes('insert');

                const operationStart = isRowOperation ? range.startRow : range.startColumn;
                const operationEnd = isRowOperation ? range.endRow : range.endColumn;
                const operationCount = operationEnd - operationStart + 1;

                const permissionListRule = this._selectionProtectionRuleModel.getSubunitRuleList(workbook.getUnitId(), worksheet.getSheetId());

                permissionListRule.forEach((rule) => {
                    const ranges = rule.ranges;
                    ranges.forEach((range) => {
                        let { startRow, endRow, startColumn, endColumn } = range;

                        if (isAddOperation) {
                            if (isRowOperation) {
                                if (operationStart <= startRow) {
                                    startRow += operationCount;
                                    endRow += operationCount;
                                }
                            } else {
                                if (operationStart <= startColumn) {
                                    startColumn += operationCount;
                                    endColumn += operationCount;
                                }
                            }
                        } else {
                            if (isRowOperation) {
                                if (operationEnd < startRow) {
                                    startRow -= operationCount;
                                    endRow -= operationCount;
                                }
                            } else {
                                if (operationEnd < startColumn) {
                                    startColumn -= operationCount;
                                    endColumn -= operationCount;
                                }
                            }
                        }

                        if (this._checkIsRightRange({ startRow, endRow, startColumn, endColumn })) {
                            range.startColumn = startColumn;
                            range.endColumn = endColumn;
                            range.startRow = startRow;
                            range.endRow = endRow;
                        }
                    });
                });

                this.disposableCollection.dispose();
                const { unitId, subUnitId } = command.params as IMoveRowsMutationParams;
                const handler = (config: EffectRefRangeParams) => {
                    return this.refRangeHandle(config, unitId, subUnitId);
                };

                const permissionRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                    return [...p, ...c.ranges];
                }, [] as IRange[]);

                permissionRanges.forEach((range) => {
                    this.disposableCollection.add(this._refRangeService.registerRefRange(range, handler, unitId, subUnitId));
                });
            }
        }));
    }

    private _checkIsRightRange(range: IRange) {
        return range.startRow <= range.endRow && range.startColumn <= range.endColumn;
    }

    private _initHeaderMovePermissionInterceptor() {
        this._headerMoveRenderController.interceptor.intercept(this._headerMoveRenderController.interceptor.getInterceptPoints().HEADER_MOVE_PERMISSION_CHECK, {
            handler: (defaultValue: Nullable<boolean>, selectionRange: IRange) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();
                const unitId = workbook.getUnitId();
                const subUnitId = worksheet.getSheetId();

                const workSheetEditPermission = this._worksheetPermissionService.getEditPermission({ unitId, subUnitId });
                if (!workSheetEditPermission) {
                    return false;
                }

                if (!selectionRange) {
                    return true;
                }

                const protectionLapRange = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                    return [...p, ...c.ranges];
                }, [] as IRange[]).filter((range) => {
                    return Rectangle.intersects(range, selectionRange);
                });

                const haveNotPermission = protectionLapRange.some((range) => {
                    const { startRow, startColumn, endRow, endColumn } = range;
                    for (let row = startRow; row <= endRow; row++) {
                        for (let col = startColumn; col <= endColumn; col++) {
                            const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                            if (permission?.Edit === false) {
                                return true;
                            }
                        }
                    }
                    return false;
                });

                return !haveNotPermission;
            },
        });
    }

    private _initHeaderResizePermissionInterceptor() {
        this._headerResizeRenderController.interceptor.intercept(this._headerResizeRenderController.interceptor.getInterceptPoints().HEADER_RESIZE_PERMISSION_CHECK, {
            handler: (defaultValue: Nullable<boolean>, rangeParams: { row?: number; col?: number }) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();

                const unitId = workbook.getUnitId();
                const subUnitId = worksheet.getSheetId();

                const workSheetEditPermission = this._worksheetPermissionService.getEditPermission({ unitId, subUnitId });
                if (!workSheetEditPermission) {
                    return false;
                }

                if (rangeParams.row) {
                    const setRowStylePermission = this._worksheetPermissionService.getSetRowStylePermission({ unitId, subUnitId });
                    if (setRowStylePermission === false) {
                        return false;
                    }
                } else if (rangeParams.col) {
                    const setColStylePermission = this._worksheetPermissionService.getSetColumnStylePermission({ unitId, subUnitId });
                    if (setColStylePermission === false) {
                        return false;
                    }
                }

                let selectionRange: Nullable<IRange>;

                if (rangeParams.row !== undefined) {
                    selectionRange = {
                        startRow: rangeParams.row,
                        endRow: rangeParams.row,
                        startColumn: 0,
                        endColumn: worksheet.getColumnCount() - 1,
                    };
                } else if (rangeParams.col !== undefined) {
                    selectionRange = {
                        startRow: 0,
                        endRow: worksheet.getRowCount() - 1,
                        startColumn: rangeParams.col,
                        endColumn: rangeParams.col,
                    };
                }

                if (!selectionRange) {
                    return true;
                }

                const protectionLapRange = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                    return [...p, ...c.ranges];
                }, [] as IRange[]).filter((range) => {
                    return Rectangle.intersects(range, selectionRange);
                });

                const haveNotPermission = protectionLapRange.some((range) => {
                    const { startRow, startColumn, endRow, endColumn } = range;
                    for (let row = startRow; row <= endRow; row++) {
                        for (let col = startColumn; col <= endColumn; col++) {
                            const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                            if (permission?.Edit === false) {
                                return true;
                            }
                        }
                    }
                    return false;
                });

                return !haveNotPermission;
            },
        });
    }

    private _initRangeFillPermissionInterceptor() {
        this._selectionRenderService.interceptor.intercept(this._selectionRenderService.interceptor.getInterceptPoints().RANGE_FILL_PERMISSION_CHECK, {
            handler: (_: Nullable<boolean>, position: { x: number; y: number; skeleton: SpreadsheetSkeleton }) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();

                const unitId = workbook.getUnitId();
                const subUnitId = worksheet.getSheetId();

                const workSheetEditPermission = this._worksheetPermissionService.getSetCellStylePermission({ unitId, subUnitId }) && this._worksheetPermissionService.getSetCellValuePermission({ unitId, subUnitId });
                if (!workSheetEditPermission) {
                    return false;
                }

                const ranges = this._selectionManagerService.getSelections()?.map((selection) => {
                    return selection.range;
                });

                const selectionRange = ranges?.find((range) => {
                    const cellPosition = position.skeleton.getCellByIndex(range.endRow, range.endColumn);
                    const missX = Math.abs(cellPosition.endX - position.x);
                    const missY = Math.abs(cellPosition.endY - position.y);
                    return missX <= 5 && missY <= 5;
                });

                if (!selectionRange) {
                    return true;
                }

                const { startRow, endRow, startColumn, endColumn } = selectionRange;

                for (let row = startRow; row <= endRow; row++) {
                    for (let col = startColumn; col <= endColumn; col++) {
                        const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                        if (permission?.Edit === false || permission?.View === false) {
                            return false;
                        }
                    }
                }
                return true;
            },
        });
    }

    private _initRangeMovePermissionInterceptor() {
        this._selectionRenderService.interceptor.intercept(this._selectionRenderService.interceptor.getInterceptPoints().RANGE_MOVE_PERMISSION_CHECK, {
            handler: (_: Nullable<boolean>, _cellInfo: null) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();
                const unitId = workbook.getUnitId();
                const subUnitId = worksheet.getSheetId();

                const workSheetEditPermission = this._worksheetPermissionService.getSetCellStylePermission({ unitId, subUnitId }) && this._worksheetPermissionService.getSetCellValuePermission({ unitId, subUnitId });
                if (!workSheetEditPermission) {
                    return false;
                }

                const ranges = this._selectionManagerService.getSelections()?.map((selection) => {
                    return selection.range;
                });

                const ruleRanges = this._selectionProtectionRuleModel.getSubunitRuleList(unitId, subUnitId).reduce((p, c) => {
                    return [...p, ...c.ranges];
                }, [] as IRange[]);

                const permissionLapRanges = ranges?.filter((range) => {
                    return ruleRanges.some((ruleRange) => {
                        return Rectangle.intersects(ruleRange, range);
                    });
                });

                const haveNotPermission = permissionLapRanges?.some((range) => {
                    const { startRow, startColumn, endRow, endColumn } = range;
                    for (let row = startRow; row <= endRow; row++) {
                        for (let col = startColumn; col <= endColumn; col++) {
                            const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                            if (permission?.Edit === false) {
                                return true;
                            }
                        }
                    }
                    return false;
                });

                return !haveNotPermission;
            },
        });
    }

    private _initSheetFindPermissionInterceptor() {
        this._sheetsFindReplaceController.interceptor.intercept(this._sheetsFindReplaceController.interceptor.getInterceptPoints().FIND_PERMISSION_CHECK, {
            handler: (_: Nullable<boolean>, _cellInfo: { row: number; col: number; unitId: string; subUnitId: string }) => {
                const { row, col, subUnitId } = _cellInfo;
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getSheetBySheetId(subUnitId);
                if (!worksheet) {
                    return false;
                }
                const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                return permission?.View !== false;
            },
        });
    }

    private _initDataValidationPermissionInterceptor() {
        this._dataValidationController.interceptor.intercept(this._dataValidationController.interceptor.getInterceptPoints().DATA_VALIDATION_PERMISSION_CHECK, {
            handler: (_: Nullable<(ISheetDataValidationRule & { disable?: boolean })[]>, rules: ISheetDataValidationRule[]) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();
                const rulesByPermissionCheck = rules.map((rule) => {
                    const { ranges } = rule;
                    const haveNotPermission = ranges?.some((range) => {
                        const { startRow, startColumn, endRow, endColumn } = range;
                        for (let row = startRow; row <= endRow; row++) {
                            for (let col = startColumn; col <= endColumn; col++) {
                                const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                                if (permission?.Edit === false || permission?.View === false) {
                                    return true;
                                }
                            }
                        }
                        return false;
                    });
                    if (haveNotPermission) {
                        return { ...rule, disable: true };
                    } else {
                        return { ...rule };
                    }
                });
                return rulesByPermissionCheck;
            },
        });
    }

    private _initConditionalFormattingPermissionInterceptor() {
        this._conditionalFormattingClearController.interceptor.intercept(this._conditionalFormattingClearController.interceptor.getInterceptPoints().CONDITIONAL_FORMATTING_PERMISSION_CHECK, {
            handler: (_: Nullable<(IConditionFormattingRule<IConditionalFormattingRuleConfig> & { disable?: boolean })[]>, rules: ((IConditionFormattingRule<IConditionalFormattingRuleConfig> & { disable?: boolean }))[]) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();
                const rulesByPermissionCheck = rules.map((rule) => {
                    const ranges = rule.ranges;
                    const haveNotPermission = ranges?.some((range) => {
                        const { startRow, startColumn, endRow, endColumn } = range;
                        for (let row = startRow; row <= endRow; row++) {
                            for (let col = startColumn; col <= endColumn; col++) {
                                const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                                if (permission?.Edit === false || permission?.View === false) {
                                    return true;
                                }
                            }
                        }
                        return false;
                    });
                    if (haveNotPermission) {
                        return { ...rule, disable: true };
                    } else {
                        return { ...rule };
                    }
                });
                return rulesByPermissionCheck;
            },
        });
    }

    private _initFreezePermissionInterceptor() {
        this._headerFreezeRenderController.interceptor.intercept(this._headerFreezeRenderController.interceptor.getInterceptPoints().FREEZE_PERMISSION_CHECK, {
            handler: (_: Nullable<boolean>, __) => {
                const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                const worksheet = workbook.getActiveSheet();
                if (!worksheet) {
                    return false;
                }
                const permission = this._worksheetPermissionService.getEditPermission({ unitId: workbook.getUnitId(), subUnitId: worksheet.getSheetId() });
                return permission;
            },
        });
    }

    private _initClipboardHook() {
        this.disposeWithMe(
            this._sheetClipboardService.addClipboardHook({
                id: SHEET_PERMISSION_PASTE_PLUGIN,
                onBeforePaste: (pasteTo) => {
                    const [ranges] = virtualizeDiscreteRanges([pasteTo.range]).ranges;
                    const startRange = this._selectionManagerService.getLast()?.range;
                    if (!startRange) {
                        return false;
                    }
                    const targetRange = {
                        startRow: startRange.startRow + ranges.startRow,
                        endRow: startRange.startRow + ranges.endRow,
                        startColumn: startRange.startColumn + ranges.startColumn,
                        endColumn: startRange.startColumn + ranges.endColumn,
                    };

                    const workbook = this._univerInstanceService.getCurrentUnitForType<Workbook>(UniverInstanceType.UNIVER_SHEET)!;
                    const worksheet = workbook.getActiveSheet();
                    const { startRow, endRow, startColumn, endColumn } = targetRange;

                    let hasPermission = true;

                    for (let row = startRow; row <= endRow; row++) {
                        for (let col = startColumn; col <= endColumn; col++) {
                            const permission = (worksheet.getCell(row, col) as (ICellDataForSheetInterceptor & { selectionProtection: ICellPermission[] }))?.selectionProtection?.[0];
                            if (permission?.Edit === false) {
                                hasPermission = false;
                                break;
                            }
                        }
                    }

                    if (!hasPermission) {
                        this._haveNotPermissionHandle(this._localService.t('permission.dialog.pasteErr'));
                    }

                    return hasPermission;
                },
            })
        );
    }
}
