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
use Exception;

// use Illuminate\Database\Eloquent\Collection;
use Illuminate\Support\Collection;
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
    
    abstract public function eachRow($row);
    
    public $fontFamily = '微软雅黑';
    public $rowHeight = 40;
    public $headerRowHeight = 40;
    public $titleRowHeight = 50;
    public $filePath;
    public $excel;
    public $headerLen;
    public $end;
    /**
     * @var Collection
     */
    public $headerData;
    
    public function setHeaderData() {
        $this->headerData = collect([]);
        if ($this->useTitle) {
            $this->headerData->push([$this->getTableTitle()]);
        }
        $this->headerData->push(array_column($this->getHeader(), 'name'));
        return $this;
    }
    
    public function __construct() {
        parent::__construct();
        $config = ['path' => $this->getTmpDir() . '/'];
        // dd($config);
        $this->excel = (new Excel($config))->fileName($this->setFilename($this->fileName)->fileName, 'Sheet1');
    }
    
    public function freezePanes(int $row = 2, int $column = 0) {
        return $this->excel->freezePanes($row, $column);        // 冻结前两行，列不冻结
    }
    
    public $useFreezePanes = true;
    
    public function beforeInsertData() {
        if ($this->useFreezePanes) {
            $this->freezePanes();        // 冻结前两行，列不冻结
        }
        return $this;
    }
    
    
    // public function storeTitle() {
    //     // title data
    //     $title = array_fill(1, $this->headerLen - 1, '');
    //     $title[0] = $this->getTableTitle();
    //     $this->headerData->push($title);
    //
    // }
    
    // 是否使用首行标题
    public $useTitle = true;
    public $titleStyle;
    
    public function setTitleStyle() {
        // title style
        $this->titleStyle =
            (new Format($this->fileHandle))
                ->fontSize(16)
                ->bold()
                ->font($this->fontFamily)
                ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                ->wrap()
                ->toResource();
    }
    
    public $headerStyle;
    
    public function setHeaderStyle() {
        // title style
        $this->headerStyle =
            (new Format($this->fileHandle))
                ->fontSize(10)
                ->font($this->fontFamily)
                ->bold()
                ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
                ->border(Format::BORDER_THIN)
                ->wrap()
                ->toResource();
    }
    
    public $globalStyle;
    
    public function setGlobalStyle() {
        // global style
        $this->globalStyle = (new Format($this->fileHandle))
            ->fontSize(10)
            ->font($this->fontFamily)
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::BORDER_THIN)
            ->wrap()
            ->toResource();
        return $this;
        // $this->excel = $this->excel->defaultFormat($globalStyle); // 默认样式
        
    }
    
    public function setColumnStyle() {
        $this->columnWidths = array_column($this->getHeader(), 'width');
        
        // 设置列宽 以及默认样式
        foreach ($this->columnWidths as $k => $columnWidth) {
            $column = $this->getColumn($k);
            $this->excel->setColumn($column . ':' . $column, $columnWidth, $this->globalStyle);
        }
    }
    
    // 开始插入数据
    public function startInsertData() {
        $this->setGlobalStyle();
        $this->setTitleStyle();
        $this->setHeaderStyle();
        $this->setColumnStyle();
        
        if ($this->exportAll) {
            // 全部导出时，分块插入数据
            $this->filePath = $this->insertHeaderData()
                                   ->chunk(function(int $times, $perPage) {
                                       return $this->exportAllHandle($times, $perPage);
                                   });
        } else {
            $this->completed = $this->data->count();
            $this->filePath = $this->insertHeaderData()
                                   ->insertNormalData($this->data);
        }
        
        return $this;
    }
    
    public function afterStore() {
    }
    
    public function initData() {
        $this->exportAll = $this->scope === Exporter::SCOPE_ALL;
        if ($this->exportAll) {
            set_time_limit(0);
            $this->setData([]);
        } else {
            $this->setData($this->buildData());
        }
        return $this->getData();
    }
    
    public $fileHandle;
    public $columnWidths;
    
    public function store() {
        $this->fileHandle = $this->excel->getHandle();
        
        $this->headerLen = count($this->getHeader());
        $this->end = $this->getColumn($this->headerLen - 1);
        
        $this->initData();
        
        $this->setHeaderData();
        
        $this->beforeInsertData();
        
        // 数据填充，输出
        $this->filePath = $this->startInsertData()->output();
        
        return $this;
    }
    
    /**
     * @Desc 全部导出的处理方法，返回分块数据集合
     * @param $times
     * @param $perPage
     * @return array|\Illuminate\Support\Collection|mixed
     * @Date 2023/6/14 17:55
     */
    public function exportAllHandle($times, $perPage) {
        return $this->buildData($times, $perPage);
        // return $this->getQuery();
    }
    
    /**
     * @Desc 设置普通行的样式
     * @return Excel
     * @Date 2023/6/14 18:12
     */
    public function setRowHeight() {
        return $this->excel->setRow($this->currentLine + 1, $this->rowHeight);
    }
    
    public function setTitleHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->titleRowHeight);                                  // title样式
    }
    
    public function setHeaderHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->headerRowHeight);
    }
    
    public $shouldDelete = false;
    public $startDataRow;     // 第三行开始数据行(0是第一行）
    public $currentLine = 0;  // 当前数据插入行
    
    public function getCurrentLine() {
        return $this->currentLine + 1;
    }
    
    public function insertHeaderData() {
        $this->startDataRow = count($this->headerData);
        foreach ($this->headerData as $row => $rowData) {
            $isHeader = true;
            if ($this->currentLine === 0 && $this->useTitle) {
                $isHeader = false;
                $this->setTitleHeight();
            } else {
                $this->setHeaderHeight();
            }
            
            foreach ($rowData as $column => $columnData) {
                $this->insertCell(
                    $this->currentLine
                    , $column
                    , $columnData
                    , null
                    , $isHeader ? $this->headerStyle : null
                );
            }
            
            $this->currentLine++;
        }
        return $this;
    }
    
    public function getIndex() {
        return $this->index;
    }
    
    /**
     * @Desc 根据序号获取rowData，分块时会被销毁
     * @param $index
     * @return mixed
     * @Date 2023/6/14 22:38
     */
    public function getRowByIndex($index) {
        return $this->data->where('index', $index)->first();
    }
    
    public function insertNormalData(Collection $data) {
        $this->data = $data;
        
        $index = $this->index;
        foreach ($this->data as $rowData) {
            if ($rowData instanceof Model) {
                $rowData->index = $index;
            }
            $index++;
        }
        foreach ($this->data as $rowData) {
            $this->setRowHeight();
            
            $rowArray = $this->eachRow($rowData);
            
            foreach ($rowArray as $column => $columnData) {
                $this->insertCell($this->currentLine, $column, $columnData);
            }
            
            $this->afterInsertOneChunkNormalData($rowData);
            
            $this->index++;
            $this->currentLine++;
        }
        
        unset($this->data, $rowArray, $rowData, $column, $columnData);
        
        return $this;
    }
    
    /**
     * @Desc 插入一个分块的数据后回调（之后分块销毁）
     * @param $rowData
     * @Date 2023/6/14 22:55
     */
    public function afterInsertOneChunkNormalData($rowData) {
    }
    
    /**
     * Insert data on the cell
     * @param int               $currentLine
     * @param int               $column
     * @param int|string|double $columnData
     * @param string|null       $format
     * @param resource|null     $formatHandle
     * @return Excel
     */
    public function insertCell(int $currentLine, int $column, $columnData, $format = null, $formatHandle = null) {
        try {
            return $this->excel->insertText($currentLine, $column, $columnData, $format, $formatHandle);
        } catch (Exception $e) {
            throw new Exception('行数为' . $this->getCurrentLine() . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
    
    /**
     * @var array 定义静态数据合并
     */
    public $mergeCellsByStaticData;
    
    public function mergeCellsAfterInsertData() {
        if ($this->useTitle) {
            return [
                ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ];
        }
        return [];
    }
    
    public function beforeOutput() {
        if (!empty($this->mergeCellsByStaticData = $this->mergeCellsAfterInsertData())) {
            foreach ($this->mergeCellsByStaticData as $i) {
                $this->excel->mergeCells($i['range'], $i['value'], $i['formatHandle'] ?? null);
            }
        }
    }
    
    public function output() {
        $this->beforeOutput();
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
            throw new Exception('超出最大列数');
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
    
    public $exportAll;
    
    public function export() {
        $this->store()->shouldDelete()->download();
    }
    
    public $max = 500000;    // 最大一次导出50万条数据
    public $chunkSize = 5000;// 分块处理 5000查一次 ，数值越大，内存占用越大
    public $completed = 0;   // 已完成
    public $debug = false;   // 调试查看导出情况
    
    public function chunk($callback = null) {
        $times = 1;
        $this->completed = 0;
        if ($this->debug) {
            $start = microtime(true);
            dump('开始内存占用：' . memory_get_peak_usage() / 1024000);
        }
        do {
            /** @var Collection $result */
            $result = $callback($times, $this->chunkSize);
            // dd($result->toArray());
            $count = count($result);
            $this->completed += $count;
            // dd($times,$result,$count);
            $this->insertNormalData($result);
            unset($result);
            if ($this->debug) {
                dump('已导出：' . $this->completed . '条，耗时' . (number_format(microtime(true) - $start, 2)) . '秒' . "-" . '内存：' . memory_get_peak_usage() / 1024000);
            }
            $times++;
        } while ($count === $this->chunkSize && $this->completed < $this->max);
        if ($this->debug) {
            dd('数据插入完成');
        }
        return $this;
    }
    
    /**
     * Get data with export query.
     * @param int $page
     * @param int $perPage
     * @return array|\Illuminate\Support\Collection|mixed
     */
    public function buildData(?int $page = null, ?int $perPage = null) {
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
