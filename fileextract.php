<?php
class DocxConversion{
    private $filename;

    public function __construct($filePath) {
        $this->filename = $filePath;
    }

    private function read_doc() {
        $fileHandle = fopen($this->filename, "r");
        $line = @fread($fileHandle, filesize($this->filename));   
        $lines = explode(chr(0x0D),$line);
        $outtext = "";
        foreach($lines as $thisline)
          {
            $pos = strpos($thisline, chr(0x00));
            if (($pos !== FALSE)||(strlen($thisline)==0))
              {
              } else {
                $outtext .= $thisline." ";
              }
          }
         $outtext = preg_replace("/[^a-zA-Z0-9\s\,\.\-\n\r\t@\/\_\(\)]/","",$outtext);
        return $outtext;
    }

    private function read_docx(){

        $striped_content = '';
        $content = '';

        $zip = zip_open($this->filename);

        if (!$zip || is_numeric($zip)) return false;

        while ($zip_entry = zip_read($zip)) {

            if (zip_entry_open($zip, $zip_entry) == FALSE) continue;

            if (zip_entry_name($zip_entry) != "word/document.xml") continue;

            $content .= zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));

            zip_entry_close($zip_entry);
        }// end while

        zip_close($zip);

        $content = str_replace('</w:r></w:p></w:tc><w:tc>', " ", $content);
        $content = str_replace('</w:r></w:p>', "\r\n", $content);
        $striped_content = strip_tags($content);

        // return $striped_content;

// $file = $striped_content;
$searchfor = ['Name','Contact No','Designation','Technologies'];
// $arr = ['name','mobile'];
// echo count($arr);
// the following line prevents the browser from parsing this as HTML.
header('Content-Type: text/plain');

// get the file contents, assuming the file to be readable (and exist)
// $contents = file_get_contents($file);
// escape special characters in the query
for($i=0;$i<=count($searchfor)-1;$i++)
{
$pattern = preg_quote($searchfor[$i], '/');

// finalise the regular expression, matching the whole line
$pattern = "/$pattern.*?\$/m";

// echo $pattern;
// search, and store all matching occurences in $matches
if(preg_match_all($pattern, $striped_content, $matches)){
$variables = array_unique($matches[0]);
$jj = implode("\n", $variables);

echo $jj;

// $my_array  = preg_split("//", $jj);

// print_r($my_array);




// echo $jj;
   // echo implode("\n",  $rgResult[0]);
}
else{
   echo "No matches found"; 
}
}

    }

 /************************excel sheet************************************/

function xlsx_to_text($input_file){
    $xml_filename = "xl/sharedStrings.xml"; //content file name
    $zip_handle = new ZipArchive;
    $output_text = "";
    if(true === $zip_handle->open($input_file)){
        if(($xml_index = $zip_handle->locateName($xml_filename)) !== false){
            $xml_datas = $zip_handle->getFromIndex($xml_index);
            $xml_handle = DOMDocument::loadXML($xml_datas, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
            $output_text = strip_tags($xml_handle->saveXML());
        }else{
            $output_text .="";
        }
        $zip_handle->close();
    }else{
    $output_text .="";
    }
    return $output_text;
}

/*************************power point files*****************************/
function pptx_to_text($input_file){
    $zip_handle = new ZipArchive;
    $output_text = "";
    if(true === $zip_handle->open($input_file)){
        $slide_number = 1; //loop through slide files
        while(($xml_index = $zip_handle->locateName("ppt/slides/slide".$slide_number.".xml")) !== false){
            $xml_datas = $zip_handle->getFromIndex($xml_index);
            $xml_handle = DOMDocument::loadXML($xml_datas, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
            $output_text .= strip_tags($xml_handle->saveXML());
            $slide_number++;
        }
        if($slide_number == 1){
            $output_text .="";
        }
        $zip_handle->close();
    }else{
    $output_text .="";
    }
    return $output_text;
}


    public function convertToText() {

        if(isset($this->filename) && !file_exists($this->filename)) {
            return "File Not exists";
        }

        $fileArray = pathinfo($this->filename);
        $file_ext  = $fileArray['extension'];
        if($file_ext == "doc" || $file_ext == "docx" || $file_ext == "xlsx" || $file_ext == "pptx")
        {
            if($file_ext == "doc") {
                return $this->read_doc();
            } elseif($file_ext == "docx") {
                return $this->read_docx();
            } elseif($file_ext == "xlsx") {
                return $this->xlsx_to_text();
            }elseif($file_ext == "pptx") {
                return $this->pptx_to_text();
            }
        } else {
            return "Invalid File Type";
        }
    }

}
$docObj = new DocxConversion("ramuexpresume.docx");
//$docObj = new DocxConversion("test.docx");
//$docObj = new DocxConversion("test.xlsx");
//$docObj = new DocxConversion("test.pptx");
echo $docText= $docObj->convertToText();
?>