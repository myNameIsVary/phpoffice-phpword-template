<?php
/*
 * @Author: ren
 * @Tool: VsCode.
 * @Date: 2020-01-10 00:13:47
 * @LastEditors  : rxm@wiki361.com
 * @LastEditTime : 2020-01-10 10:54:55
 * @FilePath: /phpoffice-phpword-template/TemplateTool.php
 */

trait TemplateTool
{
    /**
     * clear keys  ${mark-othermark} clear keys -
     * 清除key中的 '-'
     *
     * @param string $str
     * @return string
     */
    protected function cleanKey($str)
    {
        $key = rtrim(trim($str, '${'), '}');
        if (strpos($key, '-')) {
            $i = 0;
            while ($i < strlen($key)) {
                if (is_numeric($key[$i])) {
                    break;
                }
                $i++;
            }
            return substr($key, 0, $i);
        }
        return $key;
    }

    /**
     * find mark and replace   fileContent explode '$' to Array set key value clear key unset words key
     * 寻找标记 执行替换
     *
     * @param array $data &
     * @param array $imgList &
     * @param string $str fileStr
     * @param string $fileName
     * @return void
     */
    protected function findMark(&$data, &$imgList, $str, $fileName)
    {
        $strArr = explode('$', $str);
        $i      = 1;
        while ($i < count($strArr)) {
            $strArr[$i] = '$' . $strArr[$i];
            $searchStr  = substr($strArr[$i], 0, strpos($strArr[$i], '}<') + 1);
            $tmpStr     = strip_tags($searchStr);
            $key        = $this->cleanKey($tmpStr);
            if (isset($data[$key])) {
                if (is_array($data[$key])) {
                    $strArr[$i - 1] = $this->setFontColor($strArr[$i - 1] . '$', $data[$key]['color']);
                }
                $strArr[$i]   = str_replace($searchStr, is_array($data[$key]) ? $data[$key]['content'] : $data[$key], $strArr[$i]);
                $this->change = true;
                unset($data[$key]);
            } elseif (isset($imgList[$key])) {
                $this->replaceImage('${' . $key, $imgList[$key], $strArr[$i], $fileName);
                $strArr[$i]   = $this->clean('${' . $key, $strArr[$i]);
                $this->change = true;
                unset($imgList[$key]);
            }
            $i++;
        }

        if ($this->change) {
            foreach ($strArr as $v) {
                @$document .= $v;
            }
            return $document;
        }

        return false;
    }

    /**
     * set font color you need set words color FF0000 first then you can do it
     * 设置字体颜色
     *
     * @param string $str
     * @param string $color
     * @return string 
     */
    protected function setFontColor($str, $color)
    {
        $searchStr = substr($str, strrpos($str, '<w:color'));
        if ($others = substr($searchStr, strlen('<w:color w:val="FF0000"'), strpos($searchStr, '/>') - strlen('<w:color w:val="FF0000"'))) {
            $str = str_replace($searchStr, str_replace($others, '', $searchStr), $str);
        }
        $replace = '<w:color w:val="' . strtoupper(trim($color, '#')) . '"'; /*替换颜色*/
        $math    = '/\<w\:color w\:val\=\"([\S]+)\"/'; /*正则*/
        return preg_replace($math, $replace, rtrim($str, '$'));
    }

    /**
     * replace img 
     *
     * @param string $search
     * @param string $replace
     * @param string $str
     * @param string $fileName
     * @return void
     */
    protected function replaceImage($search, $replace, $str, $fileName)
    {
        if ($rid = $this->getRid($str)) {
            $fileStr = $this->getRelsFile($fileName);
            $imgName = $this->getImgName('rId' . $rid, $fileStr);
            $this->MoveImage($imgName, $replace);
        }
    }

    /**
     * clear mark '{}'
     * 清除图片标记文本
     *
     * @param string $search
     * @param string $str
     * @return string
     */
    protected function clean($search, $str)
    {
        $tmp   = substr($str, strlen($search));
        $index = strpos($tmp, '}<');
        return substr($tmp, 0, $index) . substr($tmp, $index + 1);
    }

    /**
     * delete old img  mv new img
     * 移动 替换图片
     *
     * @param string $imgName
     * @param string $replace
     * @return void
     */
    protected function MoveImage($imgName, $replace)
    {
        $this->zipClass->deleteName('word/media/' . $imgName);
        $this->zipClass->addFile($replace, 'word/media/' . $imgName);
    }

    /**
     * get img Name
     * 获取图片名字
     *
     * @param string $search
     * @param string $str
     * @return string
     */
    protected function getImgName($search, $str)
    {
        $tmpStr = substr($str, strpos($str, $search));
        $tmpStr = substr($tmpStr, 0, strpos($tmpStr, '"/>'));
        return substr($tmpStr, strrpos($tmpStr, '/') + 1);
    }

    /**
     * get Img tag rid
     * 获取图片的rid
     *
     * @param string $str
     * @return int
     */
    protected function getRid($str)
    {
        $tmpStr = substr($str, strpos($str, 'rId') + 3);
        $i      = 0;
        while ($tmpStr[$i] != '"') {
            @$rId .= $tmpStr[$i];
            $i++;
        }
        return (int) $rId ? $rId : 0;
    }

    /**
     * read dom  file zip  word/ path  fileContent
     * 读取dom文件
     *
     * @param string $file
     * @return void
     */
    protected function getFile($file)
    {
        $name = 'word/' . $file;
        return $this->zipClass->getFromName($name);
    }


    /**
     * get img file path
     * 获取图片所在引入文件
     *
     * @param string $fileName
     * @return string fileContent
     */
    protected function getRelsFile($fileName)
    {
        if ($fileName == 'main') {
            return $this->zipClass->getFromName('word/_rels/document.xml.rels');
        } else {
            return $this->zipClass->getFromName('word/_rels/' . $fileName . '.xml.rels');
        }
    }
}
