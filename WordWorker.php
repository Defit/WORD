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
    
    private $listStyle = array();
    private $sectionStyle = array();
    private $textStyle = array();
    
    private $sections = array();
    

    /**
     * Constructor
     * @param type $directory_path
     * path for working with generated .docx<br>
     * Path example : /.../sites/word_work_prj.by/tmp
     */
    public function __construct($directory_path = NULL) {
        $this->document = new  \PhpOffice\PhpWord\PhpWord();
        if(isset($directory_path)){
            $this->path = $directory_path.'/';
        }else{
            $this->path = $_SERVER['DOCUMENT_ROOT'].'/sites/word_work_prj.by/tmp/';
        }
    }
    
    /**
     * Change working directory
     * @param type $word_path 
     * change path for working with .docx to this path<br> 
     * Path example : /.../sites/word_work_prj.by/tmp
     */
    public function setPath($word_path){
        $this->path=$word_path;
    }
    
    /**
     * Set document properties
     * @param type $author
     * Author for generated document
     * @param type $title
     * Document title
     * @param type $description
     * Short document description
     * @param type $subject
     * Document topic
     * @param type $category
     * Document category
     * @param type $keyWords
     * Key words<br>
     * Example : 'cars, air'
     * @param type $company
     * Company
     * @param type $last_modifiedBy
     * Author for last modified
     */
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
    
    /**
     * Add new section in document
     * (https://phpword.readthedocs.io/en/latest/styles.html)
     * @param type $name_section
     * Name section
     * @param type $orientation
     * Orientation {'portrait', 'landscape'}
     * @param type $margin_top
     * Top margin (default = 0)
     * @param type $margin_left
     * Left margin (default = 0)
     * @param type $margin_right
     * Right margin (default = 0)
     * @param type $cols_num
     * Number of columns (default = 1)
     * @param type $start_page_number
     * Number page for section (default = 0)
     */
    public function addSection($name_section, $orientation = 'landscape',
            $margin_top = 0, $margin_left = 0, $margin_right = 0,
            $cols_num = 1, $start_page_number = 0){
        
        $sectionNewStyle = array(
         'orientation' => $orientation,
         'marginTop' => $margin_top,
         'marginLeft' => $margin_left,
             'marginRight' => $margin_right,
             'colsNum' => $cols_num,
             'pageNumberingStart' => $start_page_number,
         );
        
        $this->sections[$name_section] = $this->document->addSection($sectionNewStyle);
    }

    /**
     * Add text to specific section
     * (https://phpword.readthedocs.io/en/latest/styles.html)
     * @param type $section_name
     * Section name for adding text
     * @param type $text
     * Text for adding
     * @param type $font
     * Font name (default : 'Arial')
     * @param type $size
     * Text size (default : 14)
     * @param type $color
     * Text color (default : '000000')
     * @param type $bold
     * Bold (true/false)
     * @param type $italic
     * Italic (true/false)
     * @param type $align
     * Text align (default : 'left')
     * @param type $spaceBefore
     * Distance to paragraph
     */
    public function setTextToSection($section_name, $text, $font = 'Arial', $size = 14, $color = '000000',
            $bold = false, $italic = false, $align = 'left', $spaceBefore = 10){
        
        $fontStyle = array('name'=>$font, 'size'=>$size, 'color'=>$color, 'bold'=>$bold, 'italic'=>$italic);
        $parStyle = array('align'=>$align,'spaceBefore'=>$spaceBefore);
        
        $this->sections[$section_name]->addText(htmlspecialchars($text), $fontStyle, $parStyle);
    }
    
    /**
     * Set style for list items
     * @param INT $listType
     * TYPE SQUARE FILLED = 1<br>
     * TYPE BULLET FILLED = 3<br>
     * TYPE BULLET EMPTY = 5<br>
     * TYPE NUMBER = 7<br>
     * TYPE NUMBER NESTED = 8<br>
     * TYPE ALPHANUM = 9
     * @param type $font
     * Font name (default : 'Arial')
     * @param type $size
     * Text size (default : 14)
     * @param type $color
     * Text color (default : '000000')
     */
    public function setListStyle($listType = 3, $font = 'Arial', $size = 14, $color = '000000'){
        $this->listStyle['listType'] = $listType;
        $this->listStyle['font'] = $font;
        $this->listStyle['size'] = $size;
        $this->listStyle['color'] = $color;
    }
    
    /**
     * Add item to list
     * @param type $section_name
     * Section name for adding text
     * @param type $listItem_name
     * Text for list item
     * @param type $depth 
     * Depth for list item
     */
    public function addListItem($section_name, $listItem_name, $depth = 0){
        $fontStyle = array('name' => $this->listStyle['font'], 'size' => $this->listStyle['size'],
            'color' => $this->listStyle['color']);
        $listStyle = array('listType'=>$this->listStyle['listType']);
        
        $this->sections[$section_name]->addListItem($listItem_name,$depth,$fontStyle,$listStyle); 
    }

     /**
     * Save document
     * @param type $filename
     * File name<br>
     * Example : 'test' , '123'<br>
     * Dont writing file extension!
     * @param type $output
     * true - file downloading to client
     * false - file save to server directory
     */
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



























































