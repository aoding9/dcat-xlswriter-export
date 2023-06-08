<?php
/**
 * @Desc 导出基类
 * @User yangyang
 * @Date 2022/8/24 10:56
 */

namespace Aoding9\Dcat\Xlswriter\Export;

use Carbon\Carbon;
use Dcat\Admin\Grid\Exporter;
use Dcat\Admin\Grid\Exporters\AbstractExporter;
use Illuminate\Database\Eloquent\Collection;
use Illuminate\Database\Eloquent\Model;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

abstract class BaseExport extends AbstractExporter {
    public $header = [];
    public $fileName = '文件名';
    public $tableTitle = '表名';
    /**
     * @var Collection
     */
    public $data;

    public function getTmpDir(): string {
        $tmp = ini_get('upload_tmp_dir');

        if ($tmp !== false && file_exists($tmp)) {
            return realpath($tmp);
        }

        return realpath(sys_get_temp_dir());
    }

    public function setFilename($filename) {
        $this->fileName = $filename . Date('YmdHis') . '.xlsx';
        return $this;
    }

    public function getFilename() {
        return $this->fileName;
    }

    public function getHeader() {
        return $this->header;
    }

    public function getTableTitle() {
        return $this->tableTitle;
    }

    public function getData() {
        return $this->data;
    }

    public $index;

    public function setData($data) {
        if (!$data instanceof \Illuminate\Support\Collection) {
            $data = collect($data);
        }
        $this->data = $data;
        $this->index = 1;

        return $this;
    }

    abstract public function map($row);

    public $fontFamily = '微软雅黑';
    public $rowHeight = 40;
    public $headerRowHeight = 40;
    public $filePath;
    public $excel;
    public $headerLen;

    public function __construct() {
        parent::__construct();
        $config = ['path' => $this->getTmpDir() . '/'];
        // dd($config);
        $this->excel = (new Excel($config))->fileName($this->setFilename($this->fileName)->fileName, 'Sheet1');
    }

    public function store() {
        $fileHandle = $this->excel->getHandle();
        $format1 = new Format($fileHandle);
        $format2 = new Format($fileHandle);
        /** @var Collection $data */
        $data = $this->getData();
        $header = $this->getHeader();
        $this->headerLen = count($header);
        $columnWidths = array_column($header, 'width');
        $columnNames = array_column($header, 'name');
        // header
        $data->prepend($columnNames);
        // title
        $title = array_fill(1, $this->headerLen - 1, '');
        $title[0] = $this->getTableTitle();
        $data->prepend($title);

        // title style
        $titleStyle = $format1->fontSize(16)
                              ->bold()
                              ->font($this->fontFamily)
                              ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                              ->wrap()
                              ->toResource();

        // global style
        $globalStyle = $format2->fontSize(10)
                               ->font($this->fontFamily)
                               ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                               ->border(Format::BORDER_THIN)
                               ->wrap()
                               ->toResource();

        // 获取最后一列的列名
        $end = $this->getColumn($this->headerLen - 1);

        // 应用样式
        $this->excel = $this->excel/*->defaultFormat($globalStyle)*/// 默认样式
        ->MergeCells("A1:{$end}1", $this->getFilename())                  // 合并title单元格
        ->setRow("A1", 50, $titleStyle)                              // title样式
        ->setRow("A2", $this->headerRowHeight)                       // header样式
        ->freezePanes(2, 0);        // 冻结前两行，列不冻结
        // 设置列宽 以及默认样式

        foreach ($columnWidths as $k => $columnWidth) {
            $column = $this->getColumn($k);
            $this->excel->setColumn($column . ':' . $column, $columnWidth, $globalStyle);
        }
        // dd($data);
        // 数据填充，导出
        if ($this->isAll) {
            $this->insertData($data);
            $this->filePath = $this->chunk(function(int $times, $perPage) {
                return $this->isAllHandle($times, $perPage);
            })->output();
        } else {
            $this->filePath = $this->insertData($data)->output();
        }

        return $this;
    }

    public static function formatDate($date, $format = "m-d") {
        if ($date) {
            return (new Carbon($date))->format($format);
        }
        return null;
    }

    public function isAllHandle($times, $perPage) {
        return $this->buildData($times, $perPage);
        // return $this->getQuery();
    }

    public $shouldDelete = false;
    public $startDataRow = 2; // 第三行开始数据行(0是第一行）
    public $currentLine = 0;  // 当前数据插入行

    public function insertData($data) {
        foreach ($data as $row => $rowData) {
            // 对数据行处理
            if ($this->currentLine >= $this->startDataRow) {
                $this->excel->setRow($this->currentLine + 1, $this->rowHeight);          // 设置行高，这里的行又是从1开始的，所以+1
            }
            if ($rowData instanceof Model) {
                $rowData = $this->map($rowData);
                $this->index++;
            }
            foreach ($rowData as $column => $columnData) {
                $this->excel->insertText($this->currentLine, $column, $columnData);
            }

            $this->currentLine++;
        }
        return $this;
    }

    public function output() {
        return $this->excel->output();
    }

    public function getColumn(int $columnIndex) {
        $columnIndex++;
        $first = 64 + (int)($columnIndex / 26);
        $second = 64 + $columnIndex % 26; // 26个字母
        if ($second === 64) { // 如果余0，说明是26的倍数，末位是Z，首位暂不进位，27才进位
            $first--;
            $second = 'Z';
        } else {
            $second = chr($second);
        }
        if ($first > 90) {          //64 + 26
            throw new \Exception('超出最大列数');
        } else if ($first === 64) {
            $first = '';
        } else {
            $first = chr($first);
        }
        return $first . $second;
    }

    public function shouldDelete($v = true) {
        $this->shouldDelete = $v;
        return $this;
    }

    public function download($filePath = null) {
        if ($filePath) {
            $this->filePath = $filePath;
        }
        if ($key = request('key')) {
            $this->filePath = base64_decode($key);
        }
        response()->download($this->filePath)->deleteFileAfterSend($this->shouldDelete)->send();
        exit();
    }

    public $isAll;
    public function export() {
        $this->isAll = $this->scope === Exporter::SCOPE_ALL;
    
        if ($this->isAll) {
            set_time_limit(0);
            $this->setData([])->store();
        } else {
            $this->setData($this->buildData())->store();
        }

        $this->shouldDelete()->download();
    }

    public function chunk($callback = null) {
        $times = 1;
        $chunkSize = 5000;   // 分块处理 5000查一次数据库
        $max = 200000;       // 限制为最大一次导出20万数据
        $completed = 0;
        $debug = false;
        if ($debug) {
            $start = microtime(true);
            dump('开始：' . memory_get_peak_usage() / 1000 / 1024);
        }
        do {
            /** @var Collection $result */
            $result = $callback($times, $chunkSize);
            // dd($result->toArray());
            $count = count($result);
            $completed += $count;
            // dd($times,$result,$count);
            $this->insertData($result);
            unset($result);
            if ($debug) {
                dump($completed . '：' . (number_format(microtime(true) - $start, 2)) . "-" . memory_get_peak_usage() / 1024000);
            }
           /*else {
                dump('已导出：' . $completed . '条，耗时' . (number_format(microtime(true) - $start, 2)) . '秒');
            }*/
            $times++;
        } while ($count === $chunkSize && $completed < $max);
        if ($debug) {
            dd('数据插入完成');
        }
        return $this;
    }
    
    /**
     * Get data with export query.
     *
     * @param  int  $page
     * @param  int  $perPage
     * @return array|\Illuminate\Support\Collection|mixed
     */
    public function buildData(?int $page = null, ?int $perPage = null)
    {
        $model = $this->getGridModel();
        
        // current page
        if ($this->scope === Exporter::SCOPE_CURRENT_PAGE) {
            $page = $model->getCurrentPage();
            $perPage = $model->getPerPage();
        }
        
        $model->usePaginate(false);
        
        if ($page && $this->scope !== Exporter::SCOPE_SELECTED_ROWS) {
            $perPage = $perPage ?: $this->getChunkSize();
            
            $model->forPage($page, $perPage);
        }
        
        $array = $this->grid->processFilter();
        $model->reset();
        
        // 这里不转换为数组，直接返回模型集合
        return $this->callBuilder($array);
    }
}
