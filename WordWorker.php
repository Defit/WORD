<?php
require 'vendor/autoload.php';
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 * Description of WordWorker
 *
 * @author Defit
 */
class WordWorker {
    const DEFAULT_FONT_NAME = 'Times New Roman';
    const DEFAULT_FONT_SIZE = 14;

    private $path;
    private $document;
    private $properties;
    
    private $sections = array();
    

    
    public function __construct($directory_path = NULL) {
        $this->document = new  \PhpOffice\PhpWord\PhpWord();
        if(isset($directory_path)){
            $this->path = $directory_path.'/';
        }else{
            $this->path = $_SERVER['DOCUMENT_ROOT'].'/sites/word_work_prj.by/tmp/';
        }
    }
    
    public function setPath($word_path){
        $this->path=$word_path;
    }
    
    public function setProperties($author = 'Default',
            $title = 'New document', $description = 'Description', 
            $subject = 'Subject', $category = 'Default', $keyWords = 'key, words',
            $company = 'Default', $last_modifiedBy = 'Default'){
        
        $this->document->setDefaultFontName(self::DEFAULT_FONT_NAME);
        $this->document->setDefaultFontSize(self::DEFAULT_FONT_SIZE);
        $this->properties = $this->document->getDocInfo();  
        
        $this->properties->setCreator($author);
        $this->properties->setCompany($company);
        $this->properties->setTitle($title);
        $this->properties->setDescription($description);
        $this->properties->setCategory($category);
        $this->properties->setLastModifiedBy($last_modifiedBy);
        $this->properties->setCreated(time());
        $this->properties->setModified(time());
        $this->properties->setSubject($subject);
        $this->properties->setKeywords($keyWords); 
    }
    
    public function setSection($name_section, $orientation = 'landscape',
            $margin_top = 0, $margin_left = 0, $margin_right = 0,
            $cols_num = 1, $start_page_number = 0, $borderSize_bottom = 0,
            $borderColor_bottom = 'C0C0C0'){
        
        $sectionStyle = array(
         'orientation' => $orientation,
         'marginTop' => $margin_top,
         'marginLeft' => $margin_left,
             'marginRight' => $margin_right,
             'colsNum' => $cols_num,
             'pageNumberingStart' => $start_page_number,
             'borderBottomSize'=>$borderSize_bottom,
             'borderBottomColor'=>$borderColor_bottom
         );
        
        $this->sections[$name_section] = $this->document->addSection($sectionStyle);
    }

    public function setTextToSection($section_name, $text, $font = 'Arial', $size = 14, $color = '075776',
            $bold = false, $italic = false, $align = 'left', $spaceBefore = 10){
        
        $fontStyle = array('name'=>$font, 'size'=>$size, 'color'=>$color, 'bold'=>$bold, 'italic'=>$italic);
        $parStyle = array('align'=>$align,'spaceBefore'=>$spaceBefore);
        
        $this->sections[$section_name]->addText(htmlspecialchars($text), $fontStyle, $parStyle);
    }
    
    public function saveDoc($filename = 'default.docx', $output = true){
        $writer = \PhpOffice\PhpWord\IOFactory::createWriter($this->document,'Word2007'); 
        
        if($output){
            if (ob_get_level()) {
                ob_end_clean();
            }
            header('Content-Description: File Transfer');
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment; filename="'.$filename.'.docx'.'"');
            header('Content-Transfer-Encoding: binary');
            header('Expires: 0');
            $writer->save("php://output");
            exit;
        }else{
            $save_path = $this->path.$filename;
            $writer->save($save_path.'.docx');
        }
    }
}














































