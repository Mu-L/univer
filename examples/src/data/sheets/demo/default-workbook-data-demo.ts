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

import type { IDocumentData, IWorkbookData, LocaleType } from '@univerjs/core';
import { DataValidationErrorStyle, DataValidationOperator, DataValidationType } from '@univerjs/core';

const richTextDemo: IDocumentData = {
    id: 'd',
    body: {
        dataStream: 'Instructions: ①Project division - Fill in the specific division of labor after the project is disassembled: ②Responsible Person - Enter the responsible person\'s name here: ③Date-The specific execution time of the project (detailed to the date of a certain month), and the gray color block marks the planned real-time time of the division of labor of the project (for example, the specific execution time of [regional scene model arrangement and construction] is the 2 days marked in gray. \r\n',
        textRuns: [
            {
                st: 0,
                ed: 488,
                ts: {
                    cl: {
                        rgb: 'rgb(92,92,92)',
                    },
                },
            },
        ],
        paragraphs: [
            {
                startIndex: 489,
                paragraphStyle: {
                    spaceAbove: 10,
                    lineSpacing: 1.2,
                },
            },
        ],
    },
    documentStyle: {
        pageSize: {
            width: Number.POSITIVE_INFINITY,
            height: Number.POSITIVE_INFINITY,
        },
        marginTop: 0,
        marginBottom: 0,
        marginRight: 2,
        marginLeft: 2,
    },
};

const richTextDemo1: IDocumentData = {
    id: 'd',
    body: {
        dataStream: 'No.2824163\r\n',
        textRuns: [
            {
                st: 0,
                ed: 2,
                ts: {
                    cl: {
                        rgb: '#000',
                    },
                    fs: 20,
                },
            },
            {
                st: 3,
                ed: 10,
                ts: {
                    cl: {
                        rgb: 'rgb(255, 0, 0)',
                    },
                    fs: 20,
                },
            },
        ],
        paragraphs: [
            {
                startIndex: 10,
            },
        ],
    },
    documentStyle: {
        pageSize: {
            width: Number.POSITIVE_INFINITY,
            height: Number.POSITIVE_INFINITY,
        },
        marginTop: 0,
        marginBottom: 0,
        marginRight: 2,
        marginLeft: 2,
    },
};

const dataValidation = [
    {
        uid: 'xxx-1',
        type: DataValidationType.DECIMAL,
        ranges: [{
            startRow: 0,
            endRow: 5,
            startColumn: 0,
            endColumn: 2,
        }],
        operator: DataValidationOperator.GREATER_THAN,
        formula1: '111',
        errorStyle: DataValidationErrorStyle.STOP,
    },
    {
        uid: 'xxx-0',
        type: DataValidationType.DATE,
        ranges: [{
            startRow: 0,
            endRow: 5,
            startColumn: 3,
            endColumn: 5,
        }],
        operator: DataValidationOperator.NOT_BETWEEN,
        formula1: '2024/04/10',
        formula2: '2024/10/10',
        errorStyle: DataValidationErrorStyle.STOP,
    },
    {
        uid: 'xxx-2',
        type: DataValidationType.CHECKBOX,
        ranges: [{
            startRow: 6,
            endRow: 10,
            startColumn: 0,
            endColumn: 5,
        }],
    },
    {
        uid: 'xxx-3',
        type: DataValidationType.LIST,
        ranges: [{
            startRow: 11,
            endRow: 15,
            startColumn: 0,
            endColumn: 5,
        }],
        formula1: '1,2,3,hahaha,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18',
    },
    {
        uid: 'xxx-4',
        type: DataValidationType.CUSTOM,
        ranges: [{
            startRow: 16,
            endRow: 20,
            startColumn: 0,
            endColumn: 5,
        }],
        formula1: '=A1',
    },
    {
        uid: 'xxx-5',
        type: DataValidationType.LIST_MULTIPLE,
        ranges: [{
            startRow: 21,
            endRow: 21,
            startColumn: 0,
            endColumn: 0,
        }],
        formula1: '1,2,3,4,5,哈哈哈哈',
    },
];

export const DEFAULT_WORKBOOK_DATA_DEMO: IWorkbookData = {
    id: 'book-1',
    name: 'UniverSheet Demo',
    appVersion: '3.0.0-alpha',
    locale: 'zhCN' as LocaleType,
    sheetOrder: [],
    styles: {
        XbKcXS: {
            bg: {
                rgb: '#274fee',
            },
        },
    },
    sheets: {

        'sheet-3': {
            name: '工作表3',
            id: 'sheet-3',
            tabColor: '',
            hidden: 0,
            rowCount: 1000,
            columnCount: 20,
            zoomRatio: 1,
            freeze: {
                xSplit: 0,
                ySplit: 0,
                startRow: -1,
                startColumn: -1,
            },
            scrollTop: 0,
            scrollLeft: 0,
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
            mergeData: [],
            cellData: {},
            rowData: {},
            columnData: {},
            showGridlines: 1,
            rowHeader: {
                width: 46,
                hidden: 0,
            },
            columnHeader: {
                height: 20,
                hidden: 0,
            },
            selections: [
                'A1',
            ],
            rightToLeft: 0,
        },
    },
    resources: [
    ],
};
