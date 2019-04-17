<?php
include_once './WordWorker.php';
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

class MySQLWord extends WordWorker{
    
    public function __construct($word_worker) {
        parent::__construct();
        $this->path = $word_worker->path;
        $this->document = $word_worker->document;
        $this->properties = $word_worker->properties;
        
        $this->listStyle = $word_worker->listStyle;
        $this->sectionStyle = $word_worker->sectionStyle;
        $this->textStyle = $word_worker->textStyle;
        $this->tableStyle = $word_worker->tableStyle;
            $this->tableRowStyle = $word_worker->tableRowStyle;
            $this->tableCellStyle = $word_worker->tableCellStyle;
            
        $this->selectedTable = $word_worker->selectedTable;
        
        $this->sections = $word_worker->sections;
        $this->tables = $word_worker->tables;
    }
    
    /**
     * 
     * @param type $connection_string
     * @param type $sql_query
     * @param type $params
     * The enumeration of parameters begins with a type string, 
     * followed by the values ​​of the parameters<br>
     * Example : generateSelectTable($cs,$sql,"ssi", 'a', 'b', 5);
     * "ssi" as String, String, Integer<br>
     * There are types:<br>
     * i - Integer<br>
     * d - Double<br>
     * s - String<br>
     * b - Blob object
     */
    public function generateSelectTable($connection_string, $sql_query, ...$params){
        foreach ($connection_string as $value) {
            foreach ($value as $key => $v) {
                $connection_str_parse[$key] = $v;
            }
        }
        $connection = new mysqli($connection_str_parse['host'], $connection_str_parse['login'], 
                $connection_str_parse['password'], $connection_str_parse['db_name']);
        if ($connection->connect_error) {
            die("Ошибка: не удается подключиться: " . $connection->connect_error);
            exit();
        } 
        if(count($params) == 0){
            $result = $connection->query($sql_query);

            while ($row = mysqli_fetch_row($result)){
                $cell_number = count($row);
                $this->addTableRow();
                
                for($i = 0; $i < $cell_number; $i++){
                    $this->addTableCell($row[$i], 2500);
                }
            }
        }else{
            if (!($result = $connection->prepare($sql_query))) {
                echo "Не удалось подготовить запрос: (" . $connection->errno . ") " . $connection->error;
                exit();
            }
         
            $types = array_shift($params);
            $result->bind_param($types, ...$params);
            
            if (!$result->execute()) {
                echo "Не удалось выполнить запрос: (" . $result->errno . ") " . $result->error;
                exit();
            }

            $rows = $result->get_result()->fetch_all();
            $rows_number = count($rows);
            for($i = 0; $i < $rows_number; $i++){
                $cell_number = count($rows[$i]);
                $this->addTableRow();
                
                for($j = 0; $j < $cell_number; $j++){
                    $this->addTableCell($rows[$i][$j], 2500);
                }
            }
            $result->close();
        }
    }
}



















































































