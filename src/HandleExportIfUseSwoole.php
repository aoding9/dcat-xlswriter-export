<?php

namespace Aoding9\Dcat\Xlswriter\Export;

use Dcat\Admin\Grid;
use Dcat\Admin\Layout\Content;

/**
 * 如果通过swoole来使用dcat，由于不能调用exit方法，所以需要修改控制器index方法中的部分逻辑，以返回下载响应
 * @User yangyang
 * @Date 2023/9/16 13:03
 * @method  Grid grid
 * @method  string translation
 * @method  string title
 * @method  string[] description
 */
trait HandleExportIfUseSwoole{
    public function index(Content $content) {
        $grid = $this->grid();

        if (request($grid->exporter()->getQueryName())) {
            $grid->callBuilder();
            return $grid->handleExportRequest();
        }

        return $content
            ->translation($this->translation())
            ->title($this->title())
            ->description($this->description()['index'] ?? trans('admin.list'))
            ->body($grid);
    }


}