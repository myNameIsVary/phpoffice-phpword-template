<?php
/*
 * @Author: ren
 * @Tool: VsCode.
 * @Date: 2020-01-09 23:47:45
 * @LastEditors  : rxm@wiki361.com
 * @LastEditTime : 2020-01-10 10:53:36
 * @FilePath: /phpoffice-phpword-template/Index.php
 */

require './vendor/autoload.php';
require './WordTemplate.php';

/**
 * example to replace template
 */
class Index
{
    /**
     * img list 
     *
     * @var array
     */
    protected $imgList = [];

    /**
     * words list
     *
     * @var array
     */
    protected $dataList = [];

    public function __construct()
    {
        $this->tmpPath      = './word/tmp/ ';
        $this->resourcePath = './word/resource/';
        $this->wordTpl      = './old.docx';
        $this->imgList  = [
            'logo' => './logo.png',
        ];
        $this->dataList = [
            'abc'  => '测试',
            'bcd' => [
                'color'   => '#EE0000',
                'content' => '测试2'
            ],
        ];
        $this->format = 'word';
    }

    public function index()
    {
        if (!$filePath = $this->_replaceTpl($this->dataList)) {
            return $this->error;
        }

        /* change to pdf */
        if ($this->format == 'pdf') {
            $filePath = $this->transPDF($filePath);
        }

        header($filePath);
        die;
    }


    /**
     * replace 
     * 处理文件
     *
     * @param array $data
     * @return string
     */
    protected function _replaceTpl(array $data)
    {
        if (!is_dir($this->tmpPath)) {
            mkdir($this->tmpPath, 0777, true);
        }

        $name = $this->tmpPath . pathinfo($this->wordTpl, PATHINFO_BASENAME);

        try {
            $templateProcessor = new WordTemplate($this->wordTpl);
        } catch (\ErrorException $e) {
            @unlink($this->wordTpl);
            new Exception("wordTpl is not a read resourc");
        }

        if ($templateProcessor->replace($data, $this->imgList)) {
            $templateProcessor->saveAs($name);
            return $name;
        }
        return $this->wordTpl;
    }


    /**
     * change to pdf
     * 转换成PDF
     *
     * @param string $file
     * @return void
     */
    protected function transPDF($file)
    {
        $name = pathinfo($file, PATHINFO_FILENAME) . '.pdf';
        $handle = " libreoffice --headless --invisible --convert-to pdf " . $file . " --outdir " . $this->tmpPath;
        @exec($handle);
        @unlink($file);
        return $this->tmpPath . '/' . $name;
    }
}

(new Index())->index();
