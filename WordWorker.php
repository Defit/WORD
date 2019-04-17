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

    protected $path;
    protected $document;
    protected $properties;
    
    protected $listStyle = array();
    protected $sectionStyle = array();
    protected $textStyle = array();
    protected $tableStyle = array();
        protected $tableRowStyle = array();
        protected $tableCellStyle = array();
    
    protected $selectedTable;
        
    protected $sections = array();
    protected $tables = array();
    

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
    public function setDocProperties($author = 'Default',
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
     * Set text style
     * Example : setTextStyle(['name' => 'Arial'])
     * @param type name
     * Font name (Ex : 'Arial')
     * @param type size
     * Text size (Ex : 14)
     * @param type color
     * Text color (Ex : '000000')
     * @param type bold
     * Bold (true/false)
     * @param type italic
     * Italic (true/false)
     * @param type align
     * Text align (Ex : 'left')
     * @param type spaceBefore
     * Distance to paragraph
     */
    public function setTextStyle(...$args){
        foreach ($args as $value) {
            foreach ($value as $key => $v) {
                $this->textStyle[$key] = $v;
            }
        }
    }

    /**
     * Set default style for sections<br>
     * Example : setSectionDefaultStyle(['orientation' => 'landscape'])
     * @param type orientation
     * Orientation {'portrait', 'landscape'}
     * @param type marginTop
     * Top margin
     * @param type marginLeft
     * Left margin
     * @param type marginRight
     * Right margin
     * @param type colsNum
     * Number of columns
     * @param type pageNumberingStart
     * Number page for section
     */
    public function setSectionDefaultStyle(...$args){
        foreach ($args as $value) {
            foreach ($value as $key => $v) {
                $this->sectionStyle[$key] = $v;
            }
        }
    }

    /**
     * Add new section in document with default properties
     * @param type $name_section
     * Name section
     */
    public function addSectionDefault($name_section){
        $this->sections[$name_section] = $this->document->addSection($this->sectionStyle);
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
     */
    public function setTextToSection($section_name, $text){
        
        $fontStyle = array('name' => $this->textStyle['name'], 
            'size'=>$this->textStyle['size'], 'color'=>$this->textStyle['color'], 
            'bold'=>$this->textStyle['bold'], 'italic'=>$this->textStyle['italic']);
        $parStyle = array('align'=>$this->textStyle['align'],'spaceBefore'=>$this->textStyle['spaceBefore']);
        
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
     * Set style for table<br>
     * Example : setSectionDefaultStyle(['borderColor' => '999999'])
     * @param type borderSize
     * Border size
     * @param type borderColor
     * Border color
     * @param type afterSpacing
     * etc
     * (https://phpword.readthedocs.io/en/latest/styles.html#table)
     */
    public function setTableStyle(...$args){
        foreach ($args as $value) {
            foreach ($value as $key => $v) {
                $this->tableStyle[$key] = $v;
            }
        }
    }
    
    /**
     * Set row style
     * @param $args
     * (https://phpword.readthedocs.io/en/latest/styles.html#table)
     * Example : setTableRowStyle(['exactHeight' => -5])
     */
    public function setTableRowStyle(...$args){
        foreach ($args as $value) {
            foreach ($value as $key => $v) {
                $this->tableRowStyle[$key] = $v;
            }
        }
    }
    
    /**
     * Set cell style
     * @param $args
     * (https://phpword.readthedocs.io/en/latest/styles.html#table)
     * Example : setTableCellStyle(['borderTopSize'=>1], ['borderTopColor' =>'black'])
     */
    public function setTableCellStyle(...$args){
        foreach ($args as $value) {
            foreach ($value as $key => $v) {
                $this->tableCellStyle[$key] = $v;
            }
        }
    }
    
    /**
     * Add table to section
     * @param type $section_name
     * Section name
     * @param type $table_name
     * Table name
     */
    public function addTableToSection($section_name, $table_name){
        $this->tables[$table_name] = 
                $this->sections[$section_name]->addTable($table_name, $this->tableStyle);
    }
    
    /**
     * Select table for add cells and rows
     * @param type $table_name
     * Table name
     */
    public function selectTable($table_name){
        $this->selectedTable = $table_name;
    }
    
    /**
     * Add row
     * @param type $height_row
     * Height row
     */
    public function addTableRow($height_row = -0.5){
        $this->tables[$this->selectedTable]->addRow(
                $height_row, $this->tableRowStyle);
    }
    
    /**
     * Add cell
     * @param type $text
     * Text to cell
     * @param type $width
     * Width cell
     */
    public function addTableCell($text, $width = 2500){
        $this->tables[$this->selectedTable]->addCell(
                $width, $this->tableCellStyle)->addText(
                    $text, $this->textStyle);
    }
    
     /**
     * Save document
     * @param type $filename
     * File name<br>
     * Example : 'test' , '123'<br>
     * Dont writing file extension!
     * @param type $output
     * true - file downloading to client<br>
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





























































































