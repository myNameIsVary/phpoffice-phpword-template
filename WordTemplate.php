<?php
/*
 * @Author: ren
 * @Tool: VsCode.
 * @Date: 2020-01-09 23:48:26
 * @LastEditors  : rxm@wiki361.com
 * @LastEditTime : 2020-01-10 10:54:14
 * @FilePath: /phpoffice-phpword-template/WordTemplate.php
 */

require './TemplateTool.php';

class WordTemplate extends \PhpOffice\PhpWord\TemplateProcessor
{
    use TemplateTool;

    /**
     * document.xml
     *
     * @var [type]
     */
    protected $tempDocument;

    /**
     * change file
     *
     * @var array
     */
    protected $saveFile = [];

    /**
     * is change
     *
     * @var boolean
     */
    protected $change = false;

    /**
     * is change
     *
     * @var boolean
     */
    protected $saveChange = false;

    /**
     * rewrite join document.xml
     *
     * @param [type] $documentTemplate
     */
    public function __construct($documentTemplate)
    {
        parent::__construct($documentTemplate);
        $this->tempDocument = $this->getFile('document.xml');
    }

    /**
     * replace [words] [image] foreach all file replaced
     *
     * @param array $data
     * @param array $imgList
     * @return bool
     */
    public function replace($data, $imgList)
    {
        /* 读取 read document  */
        if ($str = $this->findMark($data, $imgList, $this->tempDocument, 'main')) {
            $this->saveFile[] = ['file' => 'main', 'xml' => $str];
            $this->saveChange = true;
        }

        /* 读取 read header */
        if ($data || $imgList) {
            foreach ($this->tempDocumentHeaders as $k => $v) {
                if ($str = $this->findMark($data, $imgList, $v, 'header' . $k)) {
                    $this->saveFile[] = ['file' => 'header', 'xml' => $str, 'index' => $k];
                    $this->saveChange = true;
                }
            }
        }

        /* 读取 read footer */
        if ($data || $imgList) {
            foreach ($this->tempDocumentFooters as $k => $v) {
                if ($str = $this->findMark($data, $imgList, $v, 'footer' . $k)) {
                    $this->saveFile[] = ['file' => 'footer', 'xml' => $str, 'index' => $k];
                    $this->saveChange = true;
                }
            }
        }
        return $this->saveChange;
    }



    /**
     * Rewrite save only replace change
     * 重写save 只替换替换过的
     *
     * @return string
     */
    public function save()
    {
        foreach ($this->saveFile as $k => $v) {
            if ($v['file'] == 'header') {
                $this->savePartWithRels($this->getHeaderName($v['index']), $v['xml']);
            } elseif ($v['file'] == 'footer') {
                $this->savePartWithRels($this->getFooterName($v['index']), $v['xml']);
            } elseif ($v['file'] == 'main') {
                $this->savePartWithRels('word/document.xml', $v['xml']);
            }
        }
        /* Close zip file */
        if (false === $this->zipClass->close()) {
            throw new Exception('Could not close zip file.'); // @codeCoverageIgnore
        }
        return $this->tempDocumentFilename;
    }
}
