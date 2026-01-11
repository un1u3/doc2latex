#include <iostream>
#include <fstream>
#include <string>
#include <vector>
#include <sstream>
#include <map>
#include <regex>
#include <zip.h>  // libzip library for reading DOCX files
#include <tinyxml2.h>  // For parsing XML

using namespace tinyxml2;

class DocxToLatexConverter {
private:
    std::string docxPath;
    std::string latexOutput;
    
    // Map for style conversions
    std::map<std::string, std::string> styleMap;
    
    void initStyleMap() {
        styleMap["b"] = "\\textbf{";
        styleMap["i"] = "\\textit{";
        styleMap["u"] = "\\underline{";
    }
    
    // Extract document.xml from docx (which is a Zzip file)
    std::string extractDocumentXml() {
        int err = 0;
        zip* archive = zip_open(docxPath.c_str(), 0, &err);
        
        if (!archive) {
            throw std::runtime_error("Cannot open DOCX file: " + docxPath);
        }
        
        // DOCX files contain document.xml in word/ directory
        zip_file* file = zip_fopen(archive, "word/document.xml", 0);
        if (!file) {
            zip_close(archive);
            throw std::runtime_error("Cannot find document.xml in DOCX");
        }
        
        // Read the file content
        std::string content;
        char buffer[1024];
        zip_int64_t bytesRead;
        
        while ((bytesRead = zip_fread(file, buffer, sizeof(buffer))) > 0) {
            content.append(buffer, bytesRead);
        }
        
        zip_fclose(file);
        zip_close(archive);
        
        return content;
    }
    
    // Parse XML and convert to LaTeX
    void parseXmlToLatex(const std::string& xmlContent) {
        XMLDocument doc;
        if (doc.Parse(xmlContent.c_str()) != XML_SUCCESS) {
            throw std::runtime_error("Failed to parse XML");
        }
        
        // Start LaTeX document
        latexOutput = "\\documentclass{article}\n";
        latexOutput += "\\usepackage[utf8]{inputenc}\n";
        latexOutput += "\\usepackage{graphicx}\n\n";
        latexOutput += "\\begin{document}\n\n";
        
        // Find the document body
        XMLElement* body = doc.FirstChildElement("w:document")
                              ->FirstChildElement("w:body");
        
        if (body) {
            processParagraphs(body);
        }
        
        latexOutput += "\n\\end{document}\n";
    }
    
    void processParagraphs(XMLElement* body) {
        for (XMLElement* para = body->FirstChildElement("w:p"); 
             para != nullptr; 
             para = para->NextSiblingElement("w:p")) {
            
            std::string paraText = processParagraph(para);
            
            if (!paraText.empty()) {
                // Check for heading styles
                if (isHeading(para)) {
                    int level = getHeadingLevel(para);
                    latexOutput += getSectionCommand(level) + "{" + paraText + "}\n\n";
                } else {
                    latexOutput += paraText + "\n\n";
                }
            }
        }
    }
    
    std::string processParagraph(XMLElement* para) {
        std::string result;
        
        // Process all runs (text segments) in paragraph
        for (XMLElement* run = para->FirstChildElement("w:r"); 
             run != nullptr; 
             run = run->NextSiblingElement("w:r")) {
            
            result += processRun(run);
        }
        
        return result;
    }
    
    std::string processRun(XMLElement* run) {
        std::string text;
        std::vector<std::string> activeStyles;
        
        // Check for formatting properties
        XMLElement* rPr = run->FirstChildElement("w:rPr");
        if (rPr) {
            if (rPr->FirstChildElement("w:b")) {
                activeStyles.push_back("b");
            }
            if (rPr->FirstChildElement("w:i")) {
                activeStyles.push_back("i");
            }
            if (rPr->FirstChildElement("w:u")) {
                activeStyles.push_back("u");
            }
        }
        
        // Get the actual text
        XMLElement* textElem = run->FirstChildElement("w:t");
        if (textElem && textElem->GetText()) {
            text = textElem->GetText();
            text = escapeLatexSpecialChars(text);
        }
        
        // Apply styles
        for (auto it = activeStyles.rbegin(); it != activeStyles.rend(); ++it) {
            text = styleMap[*it] + text + "}";
        }
        
        return text;
    }
    
    bool isHeading(XMLElement* para) {
        XMLElement* pPr = para->FirstChildElement("w:pPr");
        if (!pPr) return false;
        
        XMLElement* pStyle = pPr->FirstChildElement("w:pStyle");
        if (!pStyle) return false;
        
        const char* styleId = pStyle->Attribute("w:val");
        if (!styleId) return false;
        
        std::string style(styleId);
        return style.find("Heading") != std::string::npos || 
               style.find("heading") != std::string::npos;
    }
    
    int getHeadingLevel(XMLElement* para) {
        XMLElement* pPr = para->FirstChildElement("w:pPr");
        if (!pPr) return 1;
        
        XMLElement* pStyle = pPr->FirstChildElement("w:pStyle");
        if (!pStyle) return 1;
        
        const char* styleId = pStyle->Attribute("w:val");
        if (!styleId) return 1;
        
        std::string style(styleId);
        // Extract number from "Heading1", "Heading2", etc.
        if (style.length() > 7) {
            char lastChar = style[style.length() - 1];
            if (isdigit(lastChar)) {
                return lastChar - '0';
            }
        }
        return 1;
    }
    
    std::string getSectionCommand(int level) {
        switch(level) {
            case 1: return "\\section";
            case 2: return "\\subsection";
            case 3: return "\\subsubsection";
            default: return "\\paragraph";
        }
    }
    
    std::string escapeLatexSpecialChars(const std::string& text) {
        std::string result = text;
        
        // Escape special LaTeX characters
        std::map<std::string, std::string> replacements = {
            {"\\", "\\textbackslash{}"},
            {"&", "\\&"},
            {"%", "\\%"},
            {"$", "\\$"},
            {"#", "\\#"},
            {"_", "\\_"},
            {"{", "\\{"},
            {"}", "\\}"},
            {"~", "\\textasciitilde{}"},
            {"^", "\\textasciicircum{}"}
        };
        
        for (const auto& pair : replacements) {
            size_t pos = 0;
            while ((pos = result.find(pair.first, pos)) != std::string::npos) {
                result.replace(pos, pair.first.length(), pair.second);
                pos += pair.second.length();
            }
        }
        
        return result;
    }
    
public:
    DocxToLatexConverter(const std::string& inputPath) : docxPath(inputPath) {
        initStyleMap();
    }
    
    void convert() {
        try {
            std::cout << "Reading DOCX file: " << docxPath << std::endl;
            std::string xmlContent = extractDocumentXml();
            
            std::cout << "Parsing and converting to LaTeX..." << std::endl;
            parseXmlToLatex(xmlContent);
            
            std::cout << "Conversion completed!" << std::endl;
        } catch (const std::exception& e) {
            std::cerr << "Error: " << e.what() << std::endl;
            throw;
        }
    }
    
    void saveToFile(const std::string& outputPath) {
        std::ofstream outFile(outputPath);
        if (!outFile) {
            throw std::runtime_error("Cannot create output file: " + outputPath);
        }
        
        outFile << latexOutput;
        outFile.close();
        
        std::cout << "LaTeX file saved to: " << outputPath << std::endl;
    }
    
    std::string getLatexOutput() const {
        return latexOutput;
    }
};

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cout << "Usage: " << argv[0] << " <input.docx> [output.tex]" << std::endl;
        return 1;
    }
    
    std::string inputFile = argv[1];
    std::string outputFile = argc >= 3 ? argv[2] : "output.tex";
    
    try {
        DocxToLatexConverter converter(inputFile);
        converter.convert();
        converter.saveToFile(outputFile);
        
        std::cout << "\nSuccess! You can now compile the LaTeX file with:" << std::endl;
        std::cout << "pdflatex " << outputFile << std::endl;
        
    } catch (const std::exception& e) {
        std::cerr << "Conversion failed: " << e.what() << std::endl;
        return 1;
    }
    
    return 0;
}