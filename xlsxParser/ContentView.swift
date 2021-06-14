//
//  ContentView.swift
//  xlsxParser
//
//  Created by yujin on 2021/05/27.
//

import SwiftUI
import SSZipArchive


struct Sharedstrings {
    var sstAtri: [String:Any]  = [:]
    var sheetAtri: [String:Any]  = [:]
    var worksheetAtri: [String:Any]  = [:]
    var rowAtri: [String:Any]  = [:]
    var cAtri: [String:Any]  = [:]
    var f: [String] = [String]()
    var v: [String] = [String]()
    var si: [String] = [String]()
    var t: [String] = [String]()
}


class SharedstringsParser: XMLParser {
    // Public property to hold the result
    
    //this XMLParser class has the method parser for parsing data
    var sharedstrings: [Sharedstrings] = []
    var sharedsharedstrings: [Sharedstrings] = []
    var sharedstrings3: [Sharedstrings] = []
    var sharedstrings4: [Sharedstrings] = []
    
    private var textBuffer: String = ""
    private var nextItem: Sharedstrings? = nil
    private var nextnextItem: Sharedstrings? = nil
    private var nextItem3: Sharedstrings? = nil
    private var nextItem4: Sharedstrings? = nil
    override init(data:Data) {
        super.init(data: data)
        self.delegate = self
    }
}

struct ContentView: View {
    @State private var showDetails = false
    let c = TempC()
    init() {
        c.UnzipFile()
        startParse_sharedStrings()
        startParse_sheet()
        startParse_workbook()
        }
    var body: some View {
        Text("Hello, world!")
            .padding()
        
        Button("Show details") {
//            showDetails.toggle()
//            startParse()
        }

//        if showDetails {
//            Text("You should follow me on Twitter: @twostraws")
//                                .font(.largeTitle)
//        }
    }
    
    //Reading sharedStrings from the file
    func startParse_sharedStrings(){
        let driveURL = FileManager.default.url(forUbiquityContainerIdentifier: nil)?.appendingPathComponent("Documents").appendingPathComponent("unzipped").appendingPathComponent("xl").appendingPathComponent("sharedStrings.xml")
//        print(driveURL as Any)
        let mydata = try? Data(contentsOf: driveURL!)
        if mydata != nil {
            let parser = SharedstringsParser(data:mydata!)
                          if parser.parse() {
//                            print("sharedstrings:",parser.sharedstrings)
                            print("sstAtri",parser.sharedstrings[0].sstAtri)
                            print("si",parser.sharedstrings[0].si)
                            print("t",parser.sharedstrings[0].t)
                              //...
                          } else {
                              if let error = parser.parserError {
                                  print(error)
                              } else {
                                  print("Failed with unknown reason")
                              }
                          }
        }
    }
    
    //Reading sharedStrings from the file
    func startParse_sheet(){
        let driveURL = FileManager.default.url(forUbiquityContainerIdentifier: nil)?.appendingPathComponent("Documents").appendingPathComponent("unzipped").appendingPathComponent("xl").appendingPathComponent("worksheets").appendingPathComponent("sheet1.xml")
//        print(driveURL as Any)
        let mydata = try? Data(contentsOf: driveURL!)
        if mydata != nil {
            let parser = SharedstringsParser(data:mydata!)
                          if parser.parse() {
//                            print("sharedstrings:",parser.sharedstrings)
//                            print("worksheetAtri",parser.sharedstrings3[0].worksheetAtri)
//                            print("rowAtri",parser.sharedsharedstrings[0].rowAtri)
                            
                            for(index,element) in parser.sharedstrings.enumerated(){
                                
                                print("cAtri",parser.sharedstrings[index].cAtri)
                                print("v",parser.sharedstrings[index].v)
                                print("f",parser.sharedstrings[index].f)
                            }
                           
                            
//                            print("ROWNUM",parser.sharedsharedstrings.count)
//                            print("COLUMNNUM",parser.sharedstrings.count)

                            
                          } else {
                              if let error = parser.parserError {
                                  print(error)
                              } else {
                                  print("Failed with unknown reason")
                              }
                          }
        }
    }
    
    //Reading worksheet volume from workbook file
    func startParse_workbook(){
        let driveURL = FileManager.default.url(forUbiquityContainerIdentifier: nil)?.appendingPathComponent("Documents").appendingPathComponent("unzipped").appendingPathComponent("xl").appendingPathComponent("workbook.xml")
//        print(driveURL as Any)
        let mydata = try? Data(contentsOf: driveURL!)
        if mydata != nil {
            let parser = SharedstringsParser(data:mydata!)
                          if parser.parse() {
                            
                            for(index,element) in parser.sharedstrings4.enumerated(){
                                print("sheet",parser.sharedstrings4[index].sheetAtri)
                                print("sheet_volume",parser.sharedstrings4.count)
                         
                            }
                            
                          } else {
                              if let error = parser.parserError {
                                  print(error)
                              } else {
                                  print("Failed with unknown reason")
                              }
                          }
        }
    }
}

struct ContentView_Previews: PreviewProvider {
    static var previews: some View {
        ContentView()
    }
}


//extension ArticlesParser: XMLParserDelegate {
//    //https://developer.apple.com/forums/thread/670984


extension SharedstringsParser: XMLParserDelegate {
    //https://developer.apple.com/forums/thread/670984
    // Called when opening tag (`<elementName>`) is found
    func parser(_ parser: XMLParser, didStartElement elementName: String, namespaceURI: String?, qualifiedName qName: String?, attributes attributeDict: [String : String] = [:]) {
        switch elementName {
        case "sst":
            nextItem = Sharedstrings()
//            nextItem?.sst = attributeDict["xlmns"]!
//        print("atri",attributeDict as Any)
            nextItem?.sstAtri = attributeDict
        case "sheet":
            nextItem4 = Sharedstrings()
            nextItem4?.sheetAtri = attributeDict
        case "worksheet":
            nextItem3 = Sharedstrings()
            nextItem3?.worksheetAtri = attributeDict
        case "row":
            nextnextItem = Sharedstrings()
            nextnextItem?.rowAtri = attributeDict
        case "c":
            nextItem = Sharedstrings()
            nextItem?.cAtri = attributeDict
        case "f":
            textBuffer = ""
        case "v":
            textBuffer = ""
        case "si":
            textBuffer = ""
        case "t":
            textBuffer = ""
        default:
            print("Ignoring \(elementName)")
            break
        }
    }
    
    // Called when closing tag (`</elementName>`) is found
    func parser(_ parser: XMLParser, didEndElement elementName: String, namespaceURI: String?, qualifiedName qName: String?) {
        switch elementName {
        case "sst":
            if let item = nextItem {
                sharedstrings.append(item)
            }
        case "sheet":
            if let item = nextItem4 {
                sharedstrings4.append(item)
            }
        case "worksheet":
            if let item = nextItem3 {
                sharedstrings3.append(item)
            }
        case "row":
            if let item = nextnextItem {
                sharedsharedstrings.append(item)
            }
        case "c":
            if let item = nextItem {
                sharedstrings.append(item)
            }
        case "f":
            nextItem?.f.append(textBuffer)
        case "v":
            nextItem?.v.append(textBuffer)
        case "si":
            nextItem?.si.append(textBuffer)
        case "t":
            nextItem?.t.append(textBuffer)
        default:
            print("Ignoring \(elementName)")
            break
        }
    }
    
    // Called when a character sequence is found
    // This may be called multiple times in a single element
    func parser(_ parser: XMLParser, foundCharacters string: String) {
        textBuffer += string
    }
    
    // Called when a CDATA block is found
    func parser(_ parser: XMLParser, foundCDATA CDATABlock: Data) {
        guard let string = String(data: CDATABlock, encoding: .utf8) else {
            print("CDATA contains non-textual data, ignored")
            return
        }
        textBuffer += string
    }
    
    // For debugging
    func parser(_ parser: XMLParser, parseErrorOccurred parseError: Error) {
        print(parseError)
        print("on:", parser.lineNumber, "at:", parser.columnNumber)
    }
}

class TempC {
    func UnzipFile(){
          //do stuff remember the component Documents is needed.
        let driveURL = FileManager.default.url(forUbiquityContainerIdentifier: nil)?.appendingPathComponent("Documents").appendingPathComponent("unzipped")
    
        let filePath = FileManager.default.url(forUbiquityContainerIdentifier: nil)?.appendingPathComponent("Documents").appendingPathComponent("sheets.xlsx")
        //fruits
        print(filePath as Any)
        do {
            if !FileManager.default.fileExists(atPath: driveURL!.absoluteString) {
                    try FileManager.default.createDirectory(at: driveURL!, withIntermediateDirectories: true, attributes: nil)
            }
            
            if FileManager.default.fileExists(atPath: driveURL!.path) {
            // Unzip
            let result = SSZipArchive.unzipFile(atPath: filePath!.path,
                                                toDestination: driveURL!.path,
                                                preserveAttributes: true,
                                                overwrite: true,
                                                nestedZipLevel: 1,
                                                password: nil,
                                                error: nil,
                                                delegate: nil,
                                                progressHandler: nil,
                                                completionHandler: nil)
                                                print("result",result)
            }
        }catch{
            print(error.localizedDescription);
            return
        }
    }
}


