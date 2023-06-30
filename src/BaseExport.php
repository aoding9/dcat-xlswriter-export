<?php
/**
 * @Desc 导出基类
 * @User yangyang
 * @Date 2022/8/24 10:56
 */

namespace Aoding9\Dcat\Xlswriter\Export;

use Dcat\Admin\Grid\Exporter;
use Dcat\Admin\Grid\Exporters\AbstractExporter;
use Exception;

// use Illuminate\Database\Eloquent\Collection;
use Illuminate\Database\Eloquent\Builder;
use Illuminate\Support\Collection;
use Illuminate\Database\Eloquent\Model;
use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

abstract class BaseExport extends AbstractExporter {
    /**
     * @var array 表头
     */
    public $header = [];
    public $fileName = '文件名';
    public $tableTitle = '表名';
    /**
     * @var Collection 数据集合
     */
    public $data;
    
    /**
     * @Desc 临时保存路径
     * @return string
     * @Date 2023/6/25 19:02
     */
    public function getTmpDir(): string {
        $tmp = ini_get('upload_tmp_dir');
        
        if ($tmp !== false && file_exists($tmp)) {
            return realpath($tmp);
        }
        
        return realpath(sys_get_temp_dir());
    }
    
    /**
     * @Desc 拼接导出后文件的保存路径
     * @return string
     * @Date 2023/6/21 21:34
     */
    public function getStoreFilePath() {
        return $this->getTmpDir() . '/';
    }
    
    /**
     * @Desc 拼接完整的文件名称
     * @param $filename
     * @return $this
     * @Date 2023/6/25 19:02
     */
    public function setFilename($filename) {
        $this->fileName = $filename . Date('YmdHis') . '.xlsx';
        return $this;
    }
    
    public function getFilename() {
        return $this->fileName;
    }
    
    /**
     * @Desc 获取定义的表头
     * @return array
     * @Date 2023/6/25 19:03
     */
    public function getHeader() {
        return $this->header;
    }
    
    public function getTableTitle() {
        return $this->tableTitle;
    }
    
    public function getData() {
        return $this->data;
    }
    
    public function getChunkData() {
        return $this->chunkData;
    }
    
    /**
     * @var int 当前插入行的序号，不包括表头
     */
    public $index;
    /**
     * @var string 数据源类型，query/collection
     */
    public $dataSourceType;
    
    /**
     * @Desc 初始化数据源，判断数据源的类型
     * @param array|Collection|Builder|null $dataSource
     * @return $this
     * @Date 2023/6/21 22:02
     */
    public function initDataSource($dataSource) {
        if ($dataSource instanceof Builder) {
            $this->dataSourceType = 'query';
            $this->setQuery($dataSource);
            $dataSource = [];
        } else if (is_array($dataSource) || $dataSource instanceof Collection) {
            $this->dataSourceType = 'collection';
        } else {
            $this->dataSourceType = 'other';
        }
        // 如果是collection导入，将数据源设置到data属性
        $this->setData($dataSource);
        $this->index = 1;
        return $this;
    }
    
    /**
     * @Desc 设置数据
     * @param $data
     * @return $this
     * @Date 2023/6/25 19:03
     */
    public function setData($data) {
        if (!$data instanceof Collection) {
            $data = collect($data);
        }
        $this->data = $data;
        
        return $this;
    }
    
    /**
     * @Desc 定义数据到列的关系
     * @param $row
     * @return mixed
     * @Date 2023/6/25 19:05
     */
    abstract public function eachRow($row);
    
    public $fontFamily = '微软雅黑';  // 默认字体
    public $rowHeight = 40;           // 默认行高
    public $headerRowHeight = 40;     // 表头行高
    public $titleRowHeight = 50;      // 首行标题行高
    public $filePath;
    /**
     * @var Excel $excel xlswriter的实例
     */
    public $excel;
    /**
     * @var int 表头的最大列数
     */
    public $headerLen;
    /**
     * @var string 末尾列的字母
     */
    public $end;
    /**
     * @var Collection 表头的数据，包括标题行
     */
    public $headerData;
    
    /**
     * @Desc 设置表头数据
     * @return $this
     * @Date 2023/6/25 19:08
     */
    public function setHeaderData() {
        $this->headerData = collect([]);
        if ($this->useTitle) {
            $this->headerData->push([$this->getTableTitle()]);
        }
        $this->headerData->push(array_column($this->getHeader(), 'name'));
        return $this;
    }
    
    public $query;
    
    /**
     * 构造器
     * @param Builder|array|Collection|null $dataSource 数据源传入查询构造器、数组、集合均可，为空时需重写buildData以定义数据
     * @param null                          $time       传入时间戳以计算查询在内的导出耗时，$time =microtime(true);
     */
    public function __construct($dataSource = null, $time = null) {
        parent::__construct();
        if ($this->debug) {
            $this->time = $time ?? microtime(true);
            dump('开始内存占用：' . memory_get_peak_usage() / 1024000);
        }
        $this->init($dataSource);
    }
    
    /**
     * @var array 初始化Excel实例时的配置
     */
    public $config;
    
    public function setConfig($config = null) {
        $this->config = ['path' => $this->getStoreFilePath()];
        return $this;
    }
    
    /**
     * @Desc 设置查询构造器
     * @param $query
     * @return $this
     * @Date 2023/6/25 19:12
     */
    public function setQuery($query) {
        $this->query = $query;
        return $this;
    }
    
    /**
     * @Desc 根据fileName获取拼接后最终的文件名
     * @return string
     * @Date 2023/6/25 19:12
     */
    public function getFinalFileName() {
        return $this->setFilename($this->fileName)->getFilename();
    }
    
    /**
     * @var string 工作簿名，用于导出多个工作簿
     */
    public $sheetName = 'Sheet1';
    
    public function setSheet($name) {
        $this->sheetName = $name;
        return $this;
    }
    
    /**
     * @Desc 初始化一个新的Excel实例
     * @param $config
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:13
     */
    public function newExcel($config) {
        $this->setExcel(new Excel($config));
        return $this;
    }
    
    /**
     * @Desc 初始化导出类
     * @param mixed $dataSource
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:15
     */
    public function init($dataSource = null) {
        $this->setConfig()
             ->initDataSource($dataSource)
             ->newExcel($this->config)
            ->excel
            ->fileName($this->getFinalFileName(), $this->sheetName);
        
        return $this;
    }
    
    /**
     * @Desc 实例对象挂载到导出类
     * @param Excel $excel
     * @return $this
     * @Date 2023/6/25 19:16
     */
    public function setExcel(Excel $excel) {
        $this->excel = $excel;
        return $this;
    }
    
    public function getExcel() {
        return $this->excel;
    }
    
    /**
     * @Desc 设置表格冻结
     * @param int $row
     * @param int $column
     * @return $this
     * @Date 2023/6/25 17:59
     */
    public function freezePanes(int $row = 2, int $column = 0) {
        if ($this->useFreezePanes) {
            $this->excel->freezePanes($row, $column);        // 冻结前两行，列不冻结
        }
        return $this;
    }
    
    /**
     * @var bool 是否启用表格冻结功能
     */
    public $useFreezePanes = false;
    
    /**
     * @Desc 启用表格冻结
     * @param bool $v
     * @return $this
     * @Date 2023/6/25 19:16
     */
    public function useFreezePanes(bool $v = true) {
        $this->useFreezePanes = $v;
        return $this;
    }
    
    /**
     * @Desc 插入数据前回调
     * @return $this
     * @Date 2023/6/25 19:16
     */
    public function beforeInsertData() {
        return $this;
    }
    
    /**
     * @var bool 是否使用首行标题
     */
    public $useTitle = true;
    /**
     * @var string 首行标题的内容
     */
    public $titleStyle;
    
    /**
     * @Desc 设置标题的样式
     * @Date 2023/6/25 19:17
     */
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
        return $this;
    }
    
    /**
     * @var resource 表头样式
     */
    public $headerStyle;
    
    /**
     * @Desc 设置表头样式
     * @Date 2023/6/25 19:17
     */
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
        return $this;
    }
    
    /**
     * @var bool 是否使用全局样式代替列默认样式
     * 设置为否则使用列默认样式，会导致末尾行之后仍有边框，但是速度更快
     */
    public $useGlobalStyle = true;
    /**
     * @var resource 全局默认样式
     */
    public $globalStyle;
    
    /**
     * @Desc 设置全局默认样式
     * @return $this
     * @Date 2023/6/25 19:19
     */
    public function setGlobalStyle() {
        // global style
        $this->globalStyle = (new Format($this->fileHandle))
            ->fontSize(10)
            ->font($this->fontFamily)
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::PATTERN_NONE)
            ->wrap()
            ->toResource();
        $this->excel->defaultFormat($this->globalStyle); // 默认样式
        return $this;
    }
    
    /**
     * @var resource 数据行一般样式
     */
    public $normalStyle;
    
    public function getNormalStyle() {
        return $this->normalStyle ?: $this->normalStyle = (new Format($this->fileHandle))
            ->fontSize(10)
            ->font($this->fontFamily)
            ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER)
            ->border(Format::BORDER_THIN)
            ->wrap()
            ->toResource();
    }
    
    /**
     * @Desc 设置列默认样式
     * @Date 2023/6/25 19:20
     */
    public function setColumnStyle() {
        $this->columnWidths = array_column($this->getHeader(), 'width');
        
        // 设置列宽 以及默认样式
        foreach ($this->columnWidths as $k => $columnWidth) {
            $column = $this->getColumn($k);
            if ($this->useGlobalStyle) {
                $this->excel->setColumn($column . ':' . $column, $columnWidth);
            } else {
                $this->excel->setColumn($column . ':' . $column, $columnWidth, $this->getNormalStyle());
            }
        }
    }
    
    /**
     * @Desc 开始插入数据
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:20
     */
    public function startInsertData() {
        if ($this->useGlobalStyle) {
            $this->setGlobalStyle();
        }
        $this->setTitleStyle();
        $this->setHeaderStyle();
        $this->setColumnStyle();
        
        // 全部导出时，分块插入数据
        $this->filePath = $this->insertHeaderData()
                               ->chunk(function(int $times, $perPage) {
                                   return $this->buildData($times, $perPage);
                               });
        
        // 释放数据
        unset($this->data);
        
        return $this;
    }
    
    /**
     * @Desc 保存文件后回调
     * @Date 2023/6/25 19:21
     */
    public function afterStore() {
    }
    
    /**
     * @var resource 文件资源
     */
    public $fileHandle;
    /**
     * @var array 表头宽度数组
     */
    public $columnWidths;
    
    /**
     * @Desc 根据表头最大列数，设置末尾列名
     * @param null $end
     * @return $this
     * @Date 2023/6/25 19:24
     */
    public function setEnd($end = null) {
        $this->end = $end ?? $this->getColumn($this->headerLen - 1);
        return $this;
    }
    
    /**
     * @Desc 根据表头header的数组长度，设置最大列数
     * @param null $headerLen
     * @return $this
     * @Date 2023/6/25 19:25
     */
    public function setHeaderLen($headerLen = null) {
        $this->headerLen = $headerLen ?? count($this->getHeader());
        return $this;
    }
    
    /**
     * @Desc 挂载文件资源类
     * @param null $fileHandle
     * @return $this
     * @Date 2023/6/25 19:26
     */
    public function setFileHandle($fileHandle = null) {
        $this->fileHandle = $fileHandle ?? $this->excel->getHandle();
        return $this;
    }
    
    /**
     * @Desc 输出到文件并设置文件路径
     * @param null $filePath
     * @return $this
     * @Date 2023/6/25 19:22
     */
    public function setFilePath($filePath = null) {
        $this->filePath = $filePath ?? $this->output();
        return $this;
    }
    
    /**
     * @Desc 保存文件
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:23
     */
    public function store() {
        $this->setFileHandle()  // 设置文件处理对象
             ->freezePanes()       // 冻结前两行，列不冻结
             ->setHeaderLen() // 设置最大列数
             ->setEnd() // 设置末尾的列名
             ->setHeaderData() // 设置表头数据
             ->beforeInsertData() // 插入正式数据前回调
             ->startInsertData() // 开始插入数据
             ->afterInsertData() // 插入数据完成回调
             ->setFilePath()     // 输出文件到临时目录，并设置文件地址
             ->afterStore();
        return $this;
    }
    
    /**
     * @Desc 设置普通行的样式
     * @return Excel
     * @Date 2023/6/14 18:12
     */
    public function setRowHeight() {
        return $this->excel->setRow($this->currentLine + 1, $this->rowHeight);
    }
    
    /**
     * @Desc 设置标题行高
     * @return Excel
     * @Date 2023/6/25 19:26
     */
    public function setTitleHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->titleRowHeight);                                  // title样式
    }
    
    /**
     * @Desc 设置表头行高
     * @return Excel
     * @Date 2023/6/25 19:26
     */
    public function setHeaderHeight() {
        return $this->excel->setRow("A{$this->getCurrentLine()}", $this->headerRowHeight);
    }
    
    /**
     * @var bool 导出后是否应该删除文件
     */
    public $shouldDelete = false;
    /**
     * @var int 数据开始的行数，第一行为0
     */
    public $startDataRow;
    /**
     * @var int 当前数据插入行,第一行为0，excel显示的行数需要再此基础上+1
     */
    public $currentLine = 0;
    
    /**
     * @Desc 获取当前插入行对应到excel显示的行数
     * @return int
     * @Date 2023/6/25 19:28
     */
    public function getCurrentLine() {
        return $this->currentLine + 1;
    }
    
    /**
     * @Desc 插入表头数据
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:29
     */
    public function insertHeaderData() {
        // 设置数据开始行数
        $this->startDataRow = count($this->headerData);
        
        foreach ($this->headerData as $row => $rowData) {
            $isHeader = true;
            // 第一行为标题
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
    
    /**
     * @Desc 获取当前插入行的序号
     * @return mixed
     * @Date 2023/6/25 19:30
     */
    public function getIndex() {
        return $this->index;
    }
    
    /**
     * @Desc 根据序号获取rowData，分块时会被销毁
     * @param $index
     * @return mixed
     * @Date 2023/6/14 22:38
     */
    public function getRowInChunkByIndex($index) {
        return $this->chunkData->where('index', $index)->first();
    }
    
    /**
     * @var Collection $chunkData 分块数据
     */
    public $chunkData;
    
    /**
     * @Desc 插入分块数据到表格
     * @param Collection $data
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:32
     */
    public function insertChunkData(Collection $data) {
        $this->chunkData = $data;
        $index = $this->getIndex();
        
        // 给每行数据绑定序号
        foreach ($this->chunkData as $k => $rowData) {
            if ($rowData instanceof Model) {
                $rowData->index = $index;
            } else {
                $rowData['index'] = $index;
                $this->chunkData->put($k, $rowData);
            }
            $index++;
        }
        
        foreach ($this->chunkData as $rowData) {
            $this->setRowHeight();
            
            // 将数据传给eachRow，实现与列的对应关联
            $rowArray = $this->eachRow($rowData);
            
            // 循环该行的每个数据，插入单元格
            foreach ($rowArray as $column => $columnData) {
                $this->insertCell($this->currentLine, $column, $columnData);
            }
            
            // 行插入后回调，$this->chunkData是分块数据，绑定了index，$this->getCurrentLine()获取当前行数，$this->getRowByIndex($this->index）获取该行数据。
            $this->afterInsertEachRowInEachChunk($rowData);
            
            $this->index++;
            $this->currentLine++;
        }
        
        unset($rowArray, $column);
        
        return $this;
    }
    
    /**
     * @Desc 在分块数据插入每行后回调（到下一个分块，则上一分块被销毁）
     * @param $rowData
     * @Date 2023/6/14 22:55
     */
    public function afterInsertEachRowInEachChunk($rowData) {
    }
    
    /**
     * @Desc 根据行数和列数，得到单元格名称
     * @param int $currentLine
     * @param int $column
     * @return string
     * @Date 2023/6/25 19:37
     */
    public function getCellName(int $currentLine, int $column) {
        return $this->getColumn($column) . $currentLine;
    }
    
    public function insertCellHandle($currentLine, $column, $data, $format, $formatHandle) {
        return $this->excel->insertText($currentLine, $column, $data, $format, $formatHandle);
    }
    
    /**
     * 根据行数和列数插入数据
     * @param int               $currentLine
     * @param int               $column
     * @param int|string|double $data
     * @param string|null       $format
     * @param resource|null     $formatHandle
     * @return Excel
     */
    public function insertCell(int $currentLine, int $column, $data, ?string $format = null, $formatHandle = null) {
        try {
            if ($this->useGlobalStyle) {
                $formatHandle = $formatHandle ?? $this->getNormalStyle();
            }
            return $this->insertCellHandle($currentLine, $column, $data, $format, $formatHandle);
        } catch (Exception $e) {
            throw new Exception('行数为' . $this->getCurrentLine() . '的记录导出失败，原因：' . $e->getMessage());
        }
    }
    
    /**
     * @var array 定义静态数据合并（数据插入完成之后）
     */
    public $mergeCellsByStaticData;
    
    /**
     * @Desc 数据插入完成后合并单元格，默认合并首行标题
     * @return array|array[]
     * @Date 2023/6/25 19:38
     */
    public function mergeCellsAfterInsertData() {
        if ($this->useTitle) {
            return [
                ['range' => "A1:{$this->end}1", 'value' => $this->getTableTitle(), 'formatHandle' => $this->titleStyle],
            ];
        }
        return [];
    }
    
    /**
     * @Desc 数据插入后回调
     * @return $this
     * @Date 2023/6/25 19:39
     */
    public function afterInsertData() {
        if (!empty($this->mergeCellsByStaticData = $this->mergeCellsAfterInsertData())) {
            foreach ($this->mergeCellsByStaticData as $i) {
                $this->excel->mergeCells($i['range'], $i['value'], $i['formatHandle'] ?? null);
            }
        }
        if ($this->debug) {
            dump('触发afterInsertData-耗时' . (number_format(microtime(true) - $this->time, 2)) . '秒' . "-" . '内存：' . memory_get_peak_usage() / 1024000);
            dd('数据插入已完成');
        }
        return $this;
    }
    
    /**
     * @Desc 文件输出前回调
     * @Date 2023/6/25 19:39
     */
    public function beforeOutput() {
    }
    
    /**
     * @Desc 输出文件
     * @return string
     * @Date 2023/6/25 19:39
     */
    public function output() {
        $this->beforeOutput();
        return $this->excel->output();
    }
    
    /**
     * @var array 保存列数到列名的关系
     */
    public $columnMap = [];
    
    /**
     * @Desc 根据列数得到字母
     * 可以看做10进制转26进制，除26取余，逆序排列，把余数转成字母倒序拼接。
     * @param int $columnIndex
     * @return string
     * @Date 2023/6/15 17:51
     */
    public function getColumn(int $columnIndex) {
        if (array_key_exists($columnIndex, $this->columnMap)) {
            return $this->columnMap[$columnIndex];
        }
        
        // 由于循环条件为$divide>0，而且$columnIndex从0开始，所以+1
        $divide = $columnIndex + 1;
        $columnName = '';
        while ($divide > 0) {
            // $mod为0~25，对应26个字母，$divide初始最小为1，要-1才能得到正确的余数范围
            $mod = ($divide - 1) % 26;
            $columnName = chr(65 + $mod) . $columnName;
            $divide = (int)(($divide - $mod) / 26); // 减$mod，就是去掉末尾一位的数，除以26，相当于去掉这个数位，循环这个过程，直到取到最高位，也就是截取后的数，前面为0
        }
        return $this->columnMap[$columnIndex] = $columnName;
    }
    
    /**
     * @var array 保存列名到列数的关系
     */
    public $columnIndexMap = [];
    
    /**
     * @Desc 根据字母列名得到列数
     * @param string $columnName
     * @return float|int
     * @Date 2023/6/15 19:49
     */
    public function getColumnIndexByName(string $columnName) {
        if (array_key_exists($columnName, $this->columnIndexMap)) {
            return $this->columnIndexMap[$columnName];
        }
        // 将列名中的字母按顺序拆分成一个一个单独的字母，并进行倒序排列。
        $columnNameReverse = strrev($columnName);
        $arr = str_split($columnNameReverse);
        
        // 对每个字母进行转换，将其转换为对应的数字
        $columnIndex = 0;
        foreach ($arr as $key => $value) {
            $num = ord($value) - 64;
            $columnIndex += $num * (26 ** $key);
        }
        // 将最终计算出的列数值减去1，以得到以0为起点的列数值
        return $this->columnIndexMap[$columnName] = $columnIndex - 1;
    }
    
    /**
     * @Desc 是否在下载后删除
     * @param bool $v
     * @return $this
     * @Date 2023/6/25 19:40
     */
    public function shouldDelete($v = true) {
        $this->shouldDelete = $v;
        return $this;
    }
    
    /**
     * @Desc 执行下载
     * @param null $filePath
     * @Date 2023/6/25 19:40
     */
    public function download($filePath = null) {
        if ($filePath) {
            $this->filePath = $filePath;
        }
        response()->download($this->filePath)->deleteFileAfterSend($this->shouldDelete)->send();
        exit();
    }
    
    /**
     * @Desc 导出一条龙
     *  保存到文件-》下载-》下载后删除
     * @throws Exception
     * @Date 2023/6/25 19:41
     */
    public function export() {
        $this->store()->shouldDelete()->download();
    }
    
    /**
     * @var int 最大一次导出50万条数据
     */
    public $max = 500000;
    /**
     * @var int 分块处理 5000查一次 ，数值越大，内存占用越大
     */
    public $chunkSize = 5000;
    /**
     * @var int 已插入的数据量
     */
    public $completed = 0;
    /**
     * @var bool 是否为调试模式
     */
    public $debug = false;
    
    /**
     * 设置默认字体
     * @param string $fontFamily
     */
    public function setFontFamily(string $fontFamily) {
        $this->fontFamily = $fontFamily;
        return $this;
    }
    
    /**
     * 设置表头行高
     * @param int $headerRowHeight
     */
    public function setHeaderRowHeight(int $headerRowHeight) {
        $this->headerRowHeight = $headerRowHeight;
        return $this;
    }
    
    /**
     * 设置标题行高
     * @param int $titleRowHeight
     */
    public function setTitleRowHeight(int $titleRowHeight) {
        $this->titleRowHeight = $titleRowHeight;
        return $this;
    }
    
    /**
     * 是否使用标题行
     * @param bool $useTitle
     */
    public function setUseTitle(bool $useTitle) {
        $this->useTitle = $useTitle;
        return $this;
    }
    
    /**
     * 设置最大导出数据量
     * @param int $max
     */
    public function setMax(int $max) {
        $this->max = $max;
        return $this;
    }
    
    /**
     * 设置每个分块的数据量
     * @param int $chunkSize
     */
    public function setChunkSize(int $chunkSize) {
        $this->chunkSize = $chunkSize;
        return $this;
    }
    
    /**
     * 设置是否为调试模式
     * @param bool $debug
     */
    public function setDebug(bool $debug) {
        $this->debug = $debug;
        return $this;
    }
    
    /**
     * @var float|string 调试模式中用于计算导出耗时
     */
    public $time;
    
    /**
     * @Desc 分块处理方法
     * @param null|callable $callback
     * @return $this
     * @throws Exception
     * @Date 2023/6/25 19:45
     */
    public function chunk($callback = null) {
        $times = 1;
        $this->completed = 0;
        
        do {
            /** @var Collection $result 分块回调buildData的返回值 */
            $result = $callback($times, $this->chunkSize);
            $count = count($result);
            $this->completed += $count;
            // dd($times,$result,$count);
            
            // 插入分块数据
            $this->insertChunkData($result);
            unset($this->chunkData, $result);
            
            if ($this->debug) {
                dump('已导出：' . $this->completed . '条，耗时' . (number_format(microtime(true) - $this->time, 2)) . '秒' . "-" . '内存：' . memory_get_peak_usage() / 1024000);
            }
            $times++;
        } while ($count === $this->chunkSize && $this->completed < $this->max);
        
        return $this;
    }
    
    /**
     * 根据数据源，分块获取数据
     * @param int|null $page    第几个分块
     * @param int|null $perPage 分块大小
     * @return Collection
     * @throws Exception
     */
    public function buildData(?int $page = null, ?int $perPage = null) {
        switch ($this->dataSourceType) {
            case 'query':
                return $this->buildDataFromQuery($page, $perPage);
            case 'collection':
                return $this->buildDataFromCollection($page, $perPage);
            case 'other':
                return $this->buildDataFromOther($page, $perPage);
            default :
                throw new Exception('无效的数据源类型');
        }
    }
    
    /**
     * @Desc 用查询构造器获取分块数据
     * @param int|null $page
     * @param int|null $perPage
     * @return mixed
     * @Date 2023/6/25 19:48
     */
    public function buildDataFromQuery(?int $page = null, ?int $perPage = null) {
        return $this->query->forPage($page, $perPage)->get();
    }
    
    /**
     * @Desc 从集合获取分块数据
     * @param int|null $page
     * @param int|null $perPage
     * @return Collection
     * @Date 2023/6/25 19:49
     */
    public function buildDataFromCollection(?int $page = null, ?int $perPage = null) {
        return $this->data->forPage($page, $perPage);
    }
    
    /**
     * @Desc 从其他方式获取分块数据(dcat从grid导出查询获取数据集合，获取分块数据）
     * @param int|null $page
     * @param int|null $perPage
     * @return Collection
     * @Date 2023/6/25 19:49
     */
    public function buildDataFromOther(?int $page = null, ?int $perPage = null) {
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
    
    // /**
    //  * make代替new
    //  * @param mixed ...$params
    //  * @return $this
    //  */
    // public static function make(...$params) {
    //     return new static(...$params);
    // }
    
}
