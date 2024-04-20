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
// import '../../global';
import type { IRange, ISelectionCellWithCoord, Nullable } from '@univerjs/core';
import { BooleanNumber, ObjectMatrix, sortRules } from '@univerjs/core';

import type { BaseObject } from '../../base-object';

import { FIX_ONE_PIXEL_BLUR_OFFSET, RENDER_CLASS_TYPE } from '../../basics/const';

// import { clearLineByBorderType } from '../../basics/draw';
import { getCellPositionByIndex, getColor } from '../../basics/tools';
import type { IBoundRectNoAngle, IViewportInfo, Vector2 } from '../../basics/vector2';
import { Canvas } from '../../canvas';
import type { UniverRenderingContext } from '../../context';
import type { Engine } from '../../engine';
import type { Scene } from '../../scene';
import type { SceneViewer } from '../../scene-viewer';
import { BUFFER_EDGE_SIZE_X, BUFFER_EDGE_SIZE_Y, type Viewport } from '../../viewport';
import { Documents } from '../docs/document';
import { SpreadsheetExtensionRegistry } from '../extension';
import type { Background } from './extensions/background';
import type { Border } from './extensions/border';
import type { Font } from './extensions/font';

// import type { BorderCacheItem } from './interfaces';
import { SheetComponent } from './sheet-component';
import type { SpreadsheetSkeleton } from './sheet-skeleton';

const OBJECT_KEY = '__SHEET_EXTENSION_FONT_DOCUMENT_INSTANCE__';

export class Spreadsheet extends SheetComponent {
    private _backgroundExtension!: Background;

    private _borderExtension!: Border;

    private _fontExtension!: Font;

    private _cacheCanvas!: Canvas;
    private _cacheCanvasTop!: Canvas;
    private _cacheCanvasLeft!: Canvas;
    private _cacheCanvasLeftTop!: Canvas;
    private _cacheCanvasMap: Map<string, Canvas> = new Map();

    /**
     * 增量更新
     */
    private _refreshIncrementalState = false;

    private _forceDirty = false;
    // private _forceDirtyByViewport: Record<string, boolean> = {};
        // 无法导入
        // VIEWPORT_KEY.VIEW_MAIN: false,
        // VIEWPORT_KEY.VIEW_MAIN_TOP: false,
        // VIEWPORT_KEY.VIEW_MAIN_LEFT_TOP,



    private _overflowCacheRuntime: { [row: number]: boolean } = {};

    private _overflowCacheRuntimeRange = new ObjectMatrix<IRange>();

    private _overflowCacheRuntimeTimeout: number | NodeJS.Timeout = -1;

    private _forceDisableGridlines = false;

    private _documents: Documents = new Documents(OBJECT_KEY, undefined, {
        pageMarginLeft: 0,
        pageMarginTop: 0,
    });

    isPrinting = false;

    constructor(
        oKey: string,
        spreadsheetSkeleton?: SpreadsheetSkeleton,
        private _allowCache: boolean = true
    ) {
        super(oKey, spreadsheetSkeleton);
        if (this._allowCache) {

            // this.onIsAddedToParentObserver.add((parent) => {
            //     (parent as Scene)?.getEngine()?.onTransformChangeObservable.add(() => {
            //         this._resizeCacheCanvas();
            //     });
            //     this._resizeCacheCanvas();
            //     this._addMakeDirtyToScroll();

            //     (parent as Scene)?.getViewports().forEach(vp => vp.makeDirty());
            //     window.scene = parent;

            // });
        }

        this._initialDefaultExtension();
        this.makeDirty(true);

        window.spreadsheet = this;
    }

    get backgroundExtension() {
        return this._backgroundExtension;
    }

    get borderExtension() {
        return this._borderExtension;
    }

    get fontExtension() {
        return this._fontExtension;
    }

    override getDocuments() {
        return this._documents;
    }

    get allowCache() {
        return this._allowCache;
    }

    get forceDisableGridlines() {
        return this._forceDisableGridlines;
    }

    override draw(ctx: UniverRenderingContext, bounds?: IViewportInfo) {
        // const { parent = { scaleX: 1, scaleY: 1 } } = this;
        // const mergeData = this.getMergeData();
        // const showGridlines = this.getShowGridlines() || 1;
        const spreadsheetSkeleton = this.getSkeleton();
        if (!spreadsheetSkeleton) {
            return;
        }

        const parentScale = this.getParentScale();

        const diffRanges = this._refreshIncrementalState && bounds?.diffBounds
            ? bounds?.diffBounds?.map((bound) => spreadsheetSkeleton.getRowColumnSegmentByViewBound(bound))
            : undefined;
        const viewRanges = [spreadsheetSkeleton.getRowColumnSegmentByViewBound(bounds?.cacheBounds)];

        const extensions = this.getExtensionsByOrder();

        if((bounds?.viewPortKey === 'viewMain' || bounds?.viewPortKey === 'viewMainLeft')) {
            // console.log(bounds?.viewPortKey, 'diffRange count', diffRanges.length, 'diffY StartRow', diffRanges[0].startRow, 'diffY startColumn', diffRanges[0].startColumn , ':::::height', diffRanges[0].endRow - diffRanges[0].startRow,':::::width', diffRanges[0].endColumn - diffRanges[0].startColumn);
            console.log(bounds?.viewPortKey, ' ranges', viewRanges[0],'diffRanges', diffRanges, 'bounds', bounds?.cacheBounds, bounds?.viewBound);
        }

        const timeKey = `diffRange ${bounds?.viewPortKey}!!ext!!`;
        console.time(timeKey);
        for (const extension of extensions) {
            const timeKey = `${bounds?.viewPortKey}!!ext!!${extension.uKey}`;
            console.time(timeKey);
            // if(extension.uKey !== 'DefaultFontExtension') {
                extension.draw(ctx, parentScale, spreadsheetSkeleton, {
                    viewRanges,
                    diffRanges,
                    checkOutOfViewBound: ['viewMain','viewMainLeft', 'viewMainTop'].includes(bounds!.viewPortKey),
                });
            // }
            console.timeEnd(timeKey);
        }
        console.timeEnd(timeKey);
    }

    override isHit(coord: Vector2) {
        const oCoord = this._getInverseCoord(coord);
        const skeleton = this.getSkeleton();
        if (!skeleton) {
            return false;
        }
        const { rowHeaderWidth, columnHeaderHeight } = skeleton;
        if (oCoord.x > rowHeaderWidth && oCoord.y > columnHeaderHeight) {
            return true;
        }
        return false;
    }

    override getNoMergeCellPositionByIndex(rowIndex: number, columnIndex: number) {
        const spreadsheetSkeleton = this.getSkeleton();
        if (!spreadsheetSkeleton) {
            return;
        }
        const { rowHeightAccumulation, columnWidthAccumulation, rowHeaderWidth, columnHeaderHeight } =
            spreadsheetSkeleton;

        let { startY, endY, startX, endX } = getCellPositionByIndex(
            rowIndex,
            columnIndex,
            rowHeightAccumulation,
            columnWidthAccumulation
        );

        startY += columnHeaderHeight;
        endY += columnHeaderHeight;
        startX += rowHeaderWidth;
        endX += rowHeaderWidth;

        return {
            startY,
            endY,
            startX,
            endX,
        };
    }

    override getScrollXYByRelativeCoords(coord: Vector2) {
        const scene = this.getParent() as Scene;
        let x = 0;
        let y = 0;
        const viewPort = scene.getActiveViewportByRelativeCoord(coord);
        if (viewPort) {
            const actualX = viewPort.actualScrollX || 0;
            const actualY = viewPort.actualScrollY || 0;
            x += actualX;
            y += actualY;
        }
        return {
            x,
            y,
        };
    }


    isForceDirty(): boolean {
        return this._forceDirty;
    }

    // isForceDirtyByViewMap(viewPortKey: string){
    //     return this._forceDirtyByViewport[viewPortKey];
    // }

    /**
     * 页面初始化和调整冻结行列会用到
     * @param viewPortKey
     */
    // makeForceDirtyByViewMap(viewPortKey?: string){
        // return this._forceDirtyByViewport[viewPortKey];

        // if(!viewPortKey) {
        //     Object.keys(this._forceDirtyByViewport).forEach(key => {
        //         this._forceDirtyByViewport[key] = true;
        //         // console.log(key + ': ' + this._forceDirtyByViewport[key]);
        //     });
        // } else {
        //     this._forceDirtyByViewport[viewPortKey] = true;
        // }
    // }

    /**
     * canvas resize & zoom
     * @param state
     */
    makeForceDirty(state = true) {
        this._forceDirty = state;
    }

    setForceDisableGridlines(disabled: boolean) {
        this._forceDisableGridlines = disabled;
    }

    override getSelectionBounding(startRow: number, startColumn: number, endRow: number, endColumn: number) {
        return this.getSkeleton()?.getMergeBounding(startRow, startColumn, endRow, endColumn);
    }

    /**
     * canvas resize 并不会调用此函数
     * resize 调用的是 forceDirty
     * @param state
     * @returns
     */
    override makeDirty(state: boolean = true) {
        super.makeDirty(state);
        return this;
    }

    makeDirtyArea(dirtyBounds: IBoundRectNoAngle[]) {
        // super.makeDirty(state);
        // if(state)this.markViewPortDirty(true);
        // return this;
    }

    tickTime() {
        if(!window.lastTime) {
            window.lastTime = +new Date;
        } else {
            console.log('time', +new Date - window.lastTime);
            window.lastTime = +new Date;
        }
    }

    renderByViewport(mainCtx: UniverRenderingContext, viewportBoundsInfo: IViewportInfo, spreadsheetSkeleton: SpreadsheetSkeleton) {
        const { viewBound, cacheBounds, diffBounds, diffCacheBounds, diffX, diffY, viewPortPosition, viewPortKey, isDirty, isForceDirty, shouldCacheUpdate, cacheCanvas, leftOrigin, topOrigin } = viewportBoundsInfo;
        const { rowHeaderWidth, columnHeaderHeight } = spreadsheetSkeleton;
        const { a: scaleX = 1, d: scaleY = 1 } = mainCtx.getTransform();
        const bufferEdgeSizeX = BUFFER_EDGE_SIZE_X * scaleX / window.devicePixelRatio;
        const bufferEdgeSizeY = BUFFER_EDGE_SIZE_Y * scaleY / window.devicePixelRatio;
        // mainCtx this._scene.getEngine()?.getCanvas().getContext();


        if (viewPortKey === 'viewMain') {
            const cacheCtx = cacheCanvas.getContext();
            cacheCtx.save();

            //
            const { left, top, right, bottom } = viewPortPosition;

            const dw = right - left + rowHeaderWidth;
            const dh = bottom - top + columnHeaderHeight;


            // 没有滚动
            if (diffBounds.length === 0 || (diffX === 0 && diffY === 0) || isForceDirty) {
                //console.time('!!!viewMain_render!!!_111');
                // if (isDirty || isForceDirty || this.isForceDirty()) {
                if (isDirty || isForceDirty) {
                    cacheCtx.save();
                    cacheCtx.setTransform(1, 0, 0, 1, 0, 0);
                    cacheCanvas.clear();
                    cacheCtx.restore();
                    // cacheCtx.setTransform(sceneTrans.convert2DOMMatrix2D());

                    cacheCtx.save();


                    // 在 render() 中 mainCtx 做了这样的操作
                    // mainCtx.translateWithPrecision(rowHeaderWidth, columnHeaderHeight);
                    // 所以 cacheCtx.setTransform 已经包含了 rowHeaderWidth + scroll 距离
                    const m = mainCtx.getTransform();
                    cacheCtx.setTransform(m.a,m.b, m.c, m.d, m.e, m.f);
                    // cacheCtx.transform(1, 0, 0, 1, BUFFER_EDGE_SIZE, BUFFER_EDGE_SIZE);
                    viewportBoundsInfo.viewBound = viewportBoundsInfo.cacheBounds;
                    viewportBoundsInfo.viewPortPosition = viewportBoundsInfo.cacheViewPortPosition;
                    // cacheCtx.translate(-viewportBoundsInfo.viewBound.left + rowHeaderWidth, -viewportBoundsInfo.viewBound.top + columnHeaderHeight);

                    // 不冻结
                    // cacheCtx.translate(-viewportBoundsInfo.left + BUFFER_EDGE_SIZE+ rowHeaderWidth, -viewportBoundsInfo.viewBound.top + columnHeaderHeight);


                    // 处理相对 viewportPosition 的偏移  回到 (0, 0) 单元格位置
                    // cacheCtx.translate(-left + BUFFER_EDGE_SIZE, -top);
                    cacheCtx.translate(-leftOrigin + BUFFER_EDGE_SIZE_X, -topOrigin + BUFFER_EDGE_SIZE_Y);
                    // cacheCtx.translate(bufferEdgeSizeX, 0)

                    // extension 绘制时按照内容的左上角计算, 不考虑 rowHeaderWidth
                    this.draw(cacheCtx, viewportBoundsInfo);


                    cacheCtx.restore();
                }
                console.log('dh ', viewPortKey, dh, bottom,  top, columnHeaderHeight);
                this._applyCacheFreeze(mainCtx, cacheCanvas, bufferEdgeSizeX, bufferEdgeSizeY, dw, dh, left, top, dw, dh);
                // const pic = mainCtx.canvas.toDataURL();
                // pic;
                // //console.timeEnd('!!!viewMain_render!!!_111');
            } else {
                // diffX diffY 可以是小数
                // console.log('!!!viewMain_render!!!_222, diffBounds', diffX, diffY, isDirty)
                //console.time('!!!viewMain_render!!!_222');
                // if (this.isViewPortDirty(viewPortKey)) {
                if( isDirty ){
                    // //console.time('viewMainscroll');

                    cacheCtx.save();
                    cacheCtx.setTransform(1, 0, 0, 1, 0, 0);
                    cacheCtx.globalCompositeOperation = 'copy';
                    // cacheCtx.imageSmoothingEnabled = false;// 关闭抗锯齿  没有斜向图形不需要抗锯齿
                    cacheCtx.drawImage(cacheCanvas.getCanvasEle(), diffX * scaleX, diffY * scaleY);
                    cacheCtx.restore();

                    this._refreshIncrementalState = true;

                    // 绘制之前重设画笔位置到 spreadsheet 原点, 当没有滚动时, 这个值是 (rowHeaderWidth, colHeaderHeight)
                    cacheCtx.setTransform(mainCtx.getTransform());
                    // cacheCtx.transform(1, 0, 0, 1, BUFFER_EDGE_SIZE* scaleX,  BUFFER_EDGE_SIZE* scaleX);


                    // cacheCtx.translate(-left + BUFFER_EDGE_SIZE, -top);
                    cacheCtx.translate(-leftOrigin + BUFFER_EDGE_SIZE_X, -topOrigin + BUFFER_EDGE_SIZE_Y);


                    console.time('!!!viewMain_render_222---222');
                    if (shouldCacheUpdate) {


                        // 进入到这里的坐标, 从 sheet corner 右下角计算 也就是不算行头列头
                        // const tr = cacheCtx.getTransform();
                        // for (let index = 0; index < cacheBounds.right; ) {
                        //     cacheCtx.fillText( ''+index, index, 280)//-tr.f + 100)
                        //     cacheCtx.beginPath();
                        //     cacheCtx.moveTo(index, 280); // 将画笔移动到起点
                        //     cacheCtx.lineTo(index, 1000);     // 绘制直线到终点
                        //     cacheCtx.stroke();               // 绘制直线
                        //     index = index + 50;
                        // }

                        for (const diffBound of diffCacheBounds) {
                            cacheCtx.save();

                            const { left: diffLeft, right: diffRight, bottom: diffBottom, top: diffTop } = diffBound;
                            cacheCtx.beginPath();

                            // 再次注意!  draw 的时候 ctx.translate 单元格偏移是相对 spreadsheet content
                            // 不考虑 rowHeaderWidth
                            // 但是 diffBounds 是包括 rowHeader 信息
                            const x = diffLeft - (rowHeaderWidth) - FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            const y = diffTop - columnHeaderHeight - FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            const w = diffRight - diffLeft + (rowHeaderWidth) + FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            const h = diffBottom - diffTop + columnHeaderHeight + FIX_ONE_PIXEL_BLUR_OFFSET * 2;


                            // cacheCtx.clearRect(0, 0, 4000, 4000);
                            // console.log('xywh', x, x+w, 'trans', tr.e/tr.a, 'rs', tr.e/tr.a + x, 'diffX', diffRight - diffLeft, diffX)

                            cacheCtx.rectByPrecision(x, y, w, h);
                            // cacheCtx.fillStyle = 'rgba(220, 220, 255, 1)';
                            // cacheCtx.fill();
                            // cacheCtx.fillStyle = 'red';
                            // cacheCtx.fillText( ''+ x, x, 340)//-tr.f + 100)
                            cacheCtx.clip();


                            //@ts-ignore
                            this.draw(cacheCtx, {
                                // viewBound 是这一帧的区域
                                viewBound: viewportBoundsInfo.cacheBounds,
                                cacheBounds: viewportBoundsInfo.cacheBounds,
                                diffBounds: [diffBound],
                                // diffBounds: diffCacheBounds,
                                // diffCacheBounds,
                                diffX: viewportBoundsInfo.diffX,
                                diffY: viewportBoundsInfo.diffY,
                                viewPortPosition: viewportBoundsInfo.viewPortPosition,
                                viewPortKey: viewportBoundsInfo.viewPortKey,
                            });
                            cacheCtx.restore();
                        }
                    }
                    console.timeEnd('!!!viewMain_render_222---222');
                    this._refreshIncrementalState = false;

                }
                this._applyCacheFreeze(mainCtx, cacheCanvas, bufferEdgeSizeX, bufferEdgeSizeY, dw, dh, left, top, dw, dh);

            }
            cacheCtx.restore();
        } else if (['viewMainLeftTop', 'viewMainTop', 'viewMainLeft'].includes(viewPortKey)) {
            // const cacheCtx = this._cacheCanvasMap.get(viewPortKey)!.getContext();
            const cacheCtx = cacheCanvas.getContext();
            cacheCtx.save();
            const { left, top, right, bottom } = viewPortPosition;
            const dw = right - left + rowHeaderWidth;
            const dh = bottom - top + columnHeaderHeight;

            if (diffBounds.length === 0 || (diffX === 0 && diffY === 0) || isForceDirty) {
                // console.time(`${viewPortKey}_render!!!_111`);
                if (isDirty || isForceDirty || this.isForceDirty()) {
                    cacheCtx.save();
                    cacheCtx.setTransform(1, 0, 0, 1, 0, 0);
                    cacheCanvas.clear();
                    cacheCtx.restore();

                    cacheCtx.save();
                    // cacheCtx.setTransform(sceneTrans.convert2DOMMatrix2D());
                    cacheCtx.setTransform(mainCtx.getTransform());
                    cacheCtx.translate(-leftOrigin + BUFFER_EDGE_SIZE_X, -topOrigin + BUFFER_EDGE_SIZE_Y);


                    viewportBoundsInfo.viewBound = viewportBoundsInfo.cacheBounds;
                    this._draw(cacheCtx, viewportBoundsInfo);

                    cacheCtx.restore();
                }
                // console.log('dh ', viewPortKey, dh, bottom,  top, columnHeaderHeight);
                this._applyCacheFreeze(mainCtx, cacheCanvas, bufferEdgeSizeX, bufferEdgeSizeY, dw, dh, left, top, dw, dh);
                // console.timeEnd(`${viewPortKey}_render!!!_111`);
            } else {
                if(isDirty){
                    cacheCtx.save();
                    cacheCtx.setTransform(1, 0, 0, 1, 0, 0);
                    cacheCtx.globalCompositeOperation = 'copy';
                    // cacheCtx.imageSmoothingEnabled = false;// 关闭抗锯齿  没有斜向图形不需要抗锯齿
                    cacheCtx.drawImage(cacheCanvas.getCanvasEle(), diffX * scaleX, diffY * scaleY);
                    cacheCtx.restore();

                    this._refreshIncrementalState = true;
                    cacheCtx.setTransform(mainCtx.getTransform());
                    // cacheCtx.translate(-left + BUFFER_EDGE_SIZE, -top);
                    cacheCtx.translate(-leftOrigin + BUFFER_EDGE_SIZE_X, -topOrigin + BUFFER_EDGE_SIZE_Y);



                    // console.time(`${viewPortKey}_render!!!_222`);
                    if (shouldCacheUpdate) {
                        for (const diffBound of diffCacheBounds) {
                            const { left: diffLeft, right: diffRight, bottom: diffBottom, top: diffTop } = diffBound;
                            cacheCtx.save();
                            cacheCtx.beginPath();
                            const x = diffLeft - (rowHeaderWidth * scaleX ) + FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            const y = diffTop - columnHeaderHeight - FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            const w = diffRight - diffLeft + (rowHeaderWidth * scaleX + FIX_ONE_PIXEL_BLUR_OFFSET * 2);
                            const h = diffBottom - diffTop + columnHeaderHeight + FIX_ONE_PIXEL_BLUR_OFFSET * 2;
                            cacheCtx.rectByPrecision(x, y, w, h);

                            cacheCtx.clip();
                            // @ts-ignore
                            this._draw(cacheCtx, {
                                viewBound: viewportBoundsInfo.cacheBounds,
                                cacheBounds: viewportBoundsInfo.cacheBounds,
                                diffBounds: [diffBound],
                                diffX: viewportBoundsInfo.diffX,
                                diffY: viewportBoundsInfo.diffY,
                                viewPortPosition: viewportBoundsInfo.viewPortPosition,
                                viewPortKey: viewportBoundsInfo.viewPortKey,
                            });
                            cacheCtx.restore();
                        }
                    }
                    // console.timeEnd(`${viewPortKey}_render!!!_222`);

                    this._refreshIncrementalState = false;
                }
                this._applyCacheFreeze(mainCtx, cacheCanvas, bufferEdgeSizeX, bufferEdgeSizeY, dw, dh, left, top, dw, dh);
            }
            cacheCtx.restore();
        }

    }

    override render(mainCtx: UniverRenderingContext, bounds: IViewportInfo) {
        window.mainCtx = mainCtx;
        if (!this.visible) {
            this.makeDirty(false);
            return this;
        }

        const spreadsheetSkeleton = this.getSkeleton();

        if (!spreadsheetSkeleton) {
            return;
        }

        if(bounds.viewPortKey === 'viewMain') {
            bounds.viewBound = bounds.cacheBounds;
        }
        spreadsheetSkeleton.calculateWithoutClearingCache(bounds);

        const segment = spreadsheetSkeleton.rowColumnSegment;

        if (
            (segment.startRow === -1 && segment.endRow === -1) ||
            (segment.startColumn === -1 && segment.endColumn === -1)
        ) {
            return;
        }

        mainCtx.save();


        const { rowHeaderWidth, columnHeaderHeight } = spreadsheetSkeleton;
        mainCtx.translateWithPrecision(rowHeaderWidth, columnHeaderHeight);

        this._drawAuxiliary(mainCtx, bounds);

        const { viewPortKey } = bounds;
        if (bounds && this._allowCache === true) {
            if(['viewMain', 'viewMainLeftTop', 'viewMainTop', 'viewMainLeft'].includes(viewPortKey)) {
                this.renderByViewport(mainCtx, bounds, spreadsheetSkeleton);
            }
        } else {
            this._draw(mainCtx, bounds);
        }

        mainCtx.restore();
        return this;
    }

    /**
     * (parent as Scene)?.getEngine()?.onTransformChangeObservable.add(
     * this.onIsAddedToParentObserver.add()
     * @returns
     */
    private _resizeCacheCanvas() {
        const parentSize = this._getAncestorSize();
        if (!parentSize || this._cacheCanvas == null) {
            return;
        }
        // let { width, height } = parentSize;
        // width += BUFFER_EDGE_SIZE * 2;
        // height += BUFFER_EDGE_SIZE * 2;
        // this._cacheCanvas.setSize(width, height);
        // this._cacheCanvasTop.setSize(width, height);
        // this._cacheCanvasLeft.setSize(width, height);
        // this._cacheCanvasLeftTop.setSize(width, height);
        // this.makeDirty(true);
        // resize 后要整个重新绘制
        // render 根据 _forceDirty 才清空 cacheCanvas
        this.makeForceDirty(true);
    }

    /**
     *
     * @param mainCtx
     * @param cacheCanvas Source Image
     * @param sx
     * @param sy
     * @param sw
     * @param sh
     * @param dx
     * @param dy
     * @param dw
     * @param dh
     * @returns
     */
    protected _applyCacheFreeze(
        mainCtx: UniverRenderingContext,
        cacheCanvas: Canvas,
        sx: number = 0,
        sy: number = 0,
        sw: number = 0,
        sh: number = 0,
        dx: number = 0,
        dy: number = 0,
        dw: number = 0,
        dh: number = 0
    ) {
        if (!mainCtx) {
            return;
        }

        const pixelRatio = cacheCanvas.getPixelRatio();
        const cacheCtx = cacheCanvas.getContext();
        cacheCtx.save();
        mainCtx.save();
        mainCtx.setTransform(1, 0, 0, 1, 0, 0);
        cacheCtx.setTransform(1, 0, 0, 1, 0, 0);
        mainCtx.globalCompositeOperation = "source-over";
        mainCtx.drawImage(
            cacheCanvas.getCanvasEle(),
            sx * pixelRatio,
            sy * pixelRatio,
            sw * pixelRatio,
            sh * pixelRatio,
            dx * pixelRatio,
            dy * pixelRatio,
            dw * pixelRatio,
            dh * pixelRatio
        );
        mainCtx.restore();
        cacheCtx.restore()

        if(!document.body.contains(mainCtx.canvas)) {
            mainCtx.canvas.style.zIndex = 'BUFFER_EDGE_SIZE';
            mainCtx.canvas.style.position = 'fixed';
            mainCtx.canvas.style.background = 'lime';
            mainCtx.canvas.style.pointerEvents = 'none'; // 禁用事件响应
            mainCtx.canvas.style.border = '1px solid black'; // 设置边框样式
            mainCtx.canvas.style.transformOrigin = '30% 0%';
            mainCtx.canvas.style.transform = 'scale(0.3)';
            // document.body.appendChild(mainCtx.canvas);
        }
    }

    protected override _draw(ctx: UniverRenderingContext, bounds?: IViewportInfo) {
        this.draw(ctx, bounds);
    }

    private _getAncestorSize() {
        const parent = this._getAncestorParent();
        if (!parent) {
            return;
        }

        if (parent.classType === RENDER_CLASS_TYPE.ENGINE) {
            const mainCanvas = (parent as Engine).getCanvas();
            return {
                width: mainCanvas.getWidth(),
                height: mainCanvas.getHeight(),
            };
        }
        if (parent.classType === RENDER_CLASS_TYPE.SCENE_VIEWER) {
            return {
                width: parent.width,
                height: parent.height,
            };
        }
    }

    private _getAncestorParent(): Nullable<Engine | SceneViewer> {
        let parent: any = this.parent;
        while (parent) {
            if (parent.classType === RENDER_CLASS_TYPE.ENGINE || parent.classType === RENDER_CLASS_TYPE.SCENE_VIEWER) {
                return parent as Nullable<Engine | SceneViewer>;
            }
            parent = parent?.getParent && parent?.getParent();
        }
    }

    private _initialDefaultExtension() {
        SpreadsheetExtensionRegistry.getData()
            .sort(sortRules)
            .forEach((Extension) => {
                this.register(new Extension());
            });
        // this._borderAuxiliaryExtension = this.getExtensionByKey('DefaultBorderAuxiliaryExtension') as BorderAuxiliary;
        this._backgroundExtension = this.getExtensionByKey('DefaultBackgroundExtension') as Background;
        this._borderExtension = this.getExtensionByKey('DefaultBorderExtension') as Border;
        this._fontExtension = this.getExtensionByKey('DefaultFontExtension') as Font;
    }

    private _addMakeDirtyToScroll() {
        this._hasScrollViewportOperator(this, (viewport: Viewport) => {
            // 只有 _getHasScrollViewports() 才会进入这里  也就是 viewMain
            // console.log('!!!!!_addMakeDirtyToScroll', viewport.viewPortKey);
            viewport.onScrollBeforeObserver.add((eventData) => {
                // this.makeDirty(true);
                // eventData.viewport
                // console.log('!!_hasScrollViewportOperator', eventData.viewport?.viewPortKey);
                // this.markViewPortDirty(true, eventData.viewport?.viewPortKey);
                // this.markViewPortDirty(true);

            });
        });
    }

    private _hasScrollViewportOperator(object: BaseObject, fn: (viewPort: Viewport) => void) {
        let parent: any = object.getParent();
        while (parent) {
            if (parent.classType === RENDER_CLASS_TYPE.SCENE) {
                const viewports = parent.getViewports();
                const viewPorts = this._getHasScrollViewports(viewports);
                for (const viewport of viewPorts) {
                    if (viewport) {
                        fn(viewport);
                    }
                }
            }
            parent = parent?.getParent && parent?.getParent();
        }
    }

    private _getHasScrollViewports(viewports: Viewport[]) {
        const newViewports: Viewport[] = [];
        for (const viewport of viewports) {
            const scrollBar = viewport.getScrollBar();
            if (scrollBar) {
                newViewports.push(viewport);
            }
        }
        return newViewports;
    }

    /**
     * draw gridlines
     * @param ctx
     * @param bounds
     * @returns
     */
    private _drawAuxiliary(ctx: UniverRenderingContext, bounds?: IViewportInfo) {
        const spreadsheetSkeleton = this.getSkeleton();
        if (spreadsheetSkeleton == null) {
            return;
        }

        const { rowColumnSegment, dataMergeCache, overflowCache, stylesCache, showGridlines } = spreadsheetSkeleton;
        const { border, backgroundPositions } = stylesCache;
        const { startRow, endRow, startColumn, endColumn } = rowColumnSegment;
        if (!spreadsheetSkeleton || showGridlines === BooleanNumber.FALSE || this._forceDisableGridlines) {
            return;
        }

        const { rowHeightAccumulation, columnTotalWidth, columnWidthAccumulation, rowTotalHeight } =
            spreadsheetSkeleton;
        if (
            !rowHeightAccumulation ||
            !columnWidthAccumulation ||
            columnTotalWidth === undefined ||
            rowTotalHeight === undefined
        ) {
            return;
        }
        ctx.save();

        ctx.setLineWidthByPrecision(1);

        ctx.strokeStyle = getColor([212, 212, 212]);

        const columnWidthAccumulationLength = columnWidthAccumulation.length;
        const rowHeightAccumulationLength = rowHeightAccumulation.length;
        const EXTRA_BOUND = 0.4;
        const rowCount = endRow - startRow + 1;
        const columnCount = endColumn - startColumn + 1;
        const extraRowCount = Math.ceil(rowCount * EXTRA_BOUND);
        const extraColumnCount = Math.ceil(columnCount * EXTRA_BOUND);

        const rowStart = Math.max(Math.floor(startRow - extraRowCount), 0);
        const rowEnd = Math.min(Math.ceil(endRow + extraRowCount), rowHeightAccumulationLength - 1);
        const columnEnd = Math.min(Math.ceil(endColumn + (extraColumnCount)), columnWidthAccumulationLength - 1);
        const columnStart = Math.max(Math.floor(startColumn - (extraColumnCount)), 0);

        const startX = columnWidthAccumulation[columnStart - 1] || 0;
        const startY = rowHeightAccumulation[rowStart - 1] || 0;
        const endX = columnWidthAccumulation[columnEnd];
        const endY = rowHeightAccumulation[rowEnd];
        ctx.translateWithPrecisionRatio(FIX_ONE_PIXEL_BLUR_OFFSET, FIX_ONE_PIXEL_BLUR_OFFSET);

        ctx.beginPath();
        ctx.moveToByPrecision(startX, startY);
        ctx.lineToByPrecision(endX, startY);

        ctx.moveToByPrecision(startX, startY);
        ctx.lineToByPrecision(startX, endY);

        ctx.closePathByEnv();
        ctx.stroke();

        for (let r = rowStart; r <= rowEnd; r++) {
            if (r < 0 || r > rowHeightAccumulationLength - 1) {
                continue;
            }
            const rowEndPosition = rowHeightAccumulation[r];
            ctx.beginPath();
            ctx.moveToByPrecision(startX, rowEndPosition);
            ctx.lineToByPrecision(endX, rowEndPosition);
            ctx.closePathByEnv();
            ctx.stroke();
        }

        for (let c = columnStart; c <= columnEnd; c++) {
            if (c < 0 || c > columnWidthAccumulationLength - 1) {
                continue;
            }
            const columnEndPosition = columnWidthAccumulation[c];
            ctx.beginPath();
            ctx.moveToByPrecision(columnEndPosition, startY);
            ctx.lineToByPrecision(columnEndPosition, endY);
            ctx.closePathByEnv();
            ctx.stroke();
        }
        // console.log('xx2', scaleX, scaleY, columnTotalWidth, rowTotalHeight, rowHeightAccumulation, columnWidthAccumulation);

        // border?.forValue((rowIndex, columnIndex, borderCaches) => {
        //     if (!borderCaches) {
        //         return true;
        //     }

        //     const cellInfo = spreadsheetSkeleton.getCellByIndexWithNoHeader(rowIndex, columnIndex);

        //     let { startY, endY, startX, endX } = cellInfo;
        //     const { isMerged, isMergedMainCell, mergeInfo } = cellInfo;

        //     if (isMerged) {
        //         return true;
        //     }

        //     if (isMergedMainCell) {
        //         startY = mergeInfo.startY;
        //         endY = mergeInfo.endY;
        //         startX = mergeInfo.startX;
        //         endX = mergeInfo.endX;
        //     }

        //     if (!(mergeInfo.startRow >= rowStart && mergeInfo.endRow <= rowEnd)) {
        //         return true;
        //     }

        //     for (const key in borderCaches) {
        //         const { type } = borderCaches[key] as BorderCacheItem;

        //         clearLineByBorderType(ctx, type, { startX, startY, endX, endY });
        //     }
        // });

        // Clearing the dashed line issue caused by overlaid auxiliary lines and strokes
        // merge cell
        // this._clearRectangle(ctx, rowHeightAccumulation, columnWidthAccumulation, dataMergeCache);

        // overflow cell
        // this._clearRectangle(ctx, rowHeightAccumulation, columnWidthAccumulation, overflowCache.toNativeArray());

        // this._clearBackground(ctx, backgroundPositions);

        ctx.restore();
    }

    /**
     * Clear the guide lines within a range in the table, to make room for merged cells and overflow.
     */
    private _clearRectangle(
        ctx: UniverRenderingContext,
        rowHeightAccumulation: number[],
        columnWidthAccumulation: number[],
        dataMergeCache?: IRange[]
    ) {
        if (dataMergeCache == null) {
            return;
        }
        for (const dataCache of dataMergeCache) {
            const { startRow, endRow, startColumn, endColumn } = dataCache;

            const startY = rowHeightAccumulation[startRow - 1] || 0;
            const endY = rowHeightAccumulation[endRow] || rowHeightAccumulation[rowHeightAccumulation.length - 1];

            const startX = columnWidthAccumulation[startColumn - 1] || 0;
            const endX =
                columnWidthAccumulation[endColumn] || columnWidthAccumulation[columnWidthAccumulation.length - 1];

            ctx.clearRectByPrecision(startX, startY, endX - startX, endY - startY);

            // After ClearRect, the lines will become thinner, and the lines will be repaired below.
            ctx.beginPath();
            ctx.moveToByPrecision(startX, startY);
            ctx.lineToByPrecision(endX, startY);
            ctx.lineToByPrecision(endX, endY);
            ctx.lineToByPrecision(startX, endY);
            ctx.lineToByPrecision(startX, startY);
            ctx.stroke();
            ctx.closePath();
        }
    }

    private _clearBackground(ctx: UniverRenderingContext, backgroundPositions?: ObjectMatrix<ISelectionCellWithCoord>) {
        backgroundPositions?.forValue((row, column, cellInfo) => {
            let { startY, endY, startX, endX } = cellInfo;
            const { isMerged, isMergedMainCell, mergeInfo } = cellInfo;
            if (isMerged) {
                return true;
            }

            if (isMergedMainCell) {
                startY = mergeInfo.startY;
                endY = mergeInfo.endY;
                startX = mergeInfo.startX;
                endX = mergeInfo.endX;
            }

            ctx.clearRectForTexture(startX, startY, endX - startX + 0.5, endY - startY + 0.5);
        });
    }
}
