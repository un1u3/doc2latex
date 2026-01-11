#include <cstdio>
#include <iostream>
#include <fstream>
#include <stdexcept>
#include <string>
#include <vector>
#include <map>
#include <zip.h> //for reading ZIP arc
#include <tinyxml2.h> // for XML parsing
#include <zipconf.h>

using namespace tinyxml2;


class DocToLaxtex{

  private:
    std::string docxPath;
    std:: string latexOutput;
    std:: map<std::string, std::string> styleMap; //maps docx style to latex commands

    void initStyleMap(){
      styleMap["b"] = "\\textbf{";  //bold
      styleMap["i"] = "\\textit{";  // italic
      styleMap["u"] = "\\underline{"; //underline

    }

    std::string extractDocumentXml(){
      int err = 0;
      zip* archive = zip_open(docxPath.c_str(),0 , &err);
      if(!archive){
        throw std::runtime_error("Cannout open docx fiel"+ docxPath);

      }
      // openn documet.xml inside zip
      zip_file* file = zip_open(archive, "word/document.xml", 0);

      if(!file){
        zip_close(archive);
        thow std::runtime_error("Cannout find document.xml in docx ");

      }

      // read file contenet in chunks
      std::string content;
      char buffer[1024];
      zip_int64_t bytesRead;

      while((bytesRead = zip_fread(file, buffer,sizeof(buffer)))> 0){
        content.append(buffer, bytesRead);
      }

      zip_fclose(file);
      zip_close(archive);
      return content;
    }
}

int main(){



}
