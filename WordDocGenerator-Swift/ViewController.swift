//
//  ViewController.swift
//  WordDocGenerator-Swift
//
//  Created by AmitB on 02/07/15.
//  Copyright © 2015 Company_Name. All rights reserved.
//

import UIKit
import Foundation

class ViewController: UIViewController, WordDocumentGeneratorDelegate, UIDocumentInteractionControllerDelegate {
    
    var productArray = [Dictionary<String,AnyObject>]()
    var imagePaths : [String] = [String]()
    var wordDocDictionary : String?
    
    override func viewDidLoad() {
        super.viewDidLoad()
        // Do any additional setup after loading the view, typically from a nib.
        let imagePathArray : [NSURL] = NSBundle.mainBundle().URLsForResourcesWithExtension("jpeg", subdirectory: nil)!
        var k : Int = 0
        
        do {
            try NSFileManager.defaultManager().createDirectoryAtPath(String(format: "%@/savedimages", getDocumentsDirectory()), withIntermediateDirectories: true, attributes: nil)
            
            for imageUrl in imagePathArray
            {
                let newPath : String = String(format: "%@/savedimages/item%i.jpeg", getDocumentsDirectory(), k)
                
                let tempImage : UIImage = UIImage(named: (imageUrl as NSURL).relativeString!)!
                UIImageJPEGRepresentation(tempImage, 1.0)?.writeToFile(newPath, atomically: true)
                self.productArray.append(["name":"Item1", "price":"100", "date":"01/07/2015", "image":newPath])
                
                k++
            }
            
        }catch let error as NSError
        {
            print("Error while copying image files \(error)")
        }
    }

    override func didReceiveMemoryWarning() {
        super.didReceiveMemoryWarning()
        // Dispose of any resources that can be recreated.
    }

    @IBAction func saveToDocument(sender: AnyObject) {
        let dateFormatter : NSDateFormatter = NSDateFormatter()
        dateFormatter.dateFormat = "MM/dd/yyyy HH:mm:ss a"
        self.wordDocDictionary = dateFormatter.stringFromDate(NSDate())
        
        self.imageDownloadComplete()
    }
    
    func imageDownloadComplete()
    {
        /*
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:mv="urn:schemas-microsoft-com:mac:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body>
        */
    
        var allProducrString : String = ""
        
        var docString = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        docString = docString.stringByAppendingString("<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mo=\"http://schemas.microsoft.com/office/mac/office/2008/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wp14\"><w:body>")
        
        do {
        
            let docXML = try String(contentsOfFile: (NSBundle.mainBundle().resourcePath?.stringByAppendingPathComponent("document.xml"))!, encoding: NSUTF8StringEncoding)
            let imageXML = try String(contentsOfFile: (NSBundle.mainBundle().resourcePath?.stringByAppendingPathComponent("image.xml"))!, encoding: NSUTF8StringEncoding)
        
            for productNode in self.productArray
            {
                let productName : String = (productNode["name"] != nil) ? productNode["name"] as! String : ""
                let price : AnyObject = productNode["price"]!
                var priceLabelText = ""
                
                priceLabelText = String(format: "%.2f", price.floatValue)
                
                let dateCreated : String = productNode["date"] as! String
                var productString : String = String(format:docXML, productName, priceLabelText, dateCreated, "")
                
                productString = self.writeImagesInto(productString, images:[productNode["image"] as! String], imageXML: imageXML)
                allProducrString = allProducrString.stringByAppendingString(productString)
                
            }
            
            docString = docString.stringByAppendingString(allProducrString)
            
            docString = docString.stringByAppendingString("<w:sectPr w:rsidR=\"007E7D93\" w:rsidSect=\"00F220F0\"><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/><w:cols w:space=\"720\"/><w:docGrid w:linePitch=\"360\"/></w:sectPr>")
            docString = docString.stringByAppendingString("</w:body></w:document>")
            
            WordDocumentGenerator.sharedInstance.generateWordDocumentWithXML(docString, imagePaths: self.imagePaths, destinationDirectory: (self.wordDocDictionary?.stringByReplacingOccurrencesOfString(":", withString: "").stringByReplacingOccurrencesOfString("/", withString: "-").stringByReplacingOccurrencesOfString(" ", withString: ""))!, delegate: self)
            
        }catch let error as NSError
        {
            print("Error while processing into Controller for document \(error)")
        }
    }
    
    func writeImagesInto(var docString: String, images:[String], imageXML:String) -> String
    {
        if images.count > 0
        {
            for imageURL in images {
                docString = docString.stringByAppendingFormat(imageXML, imageURL)
                self.imagePaths.append(imageURL)
            }
        }
        
        return docString
    }
    
    func tableView(tableView: UITableView, numberOfRowsInSection section: Int) -> Int
    {
        return self.productArray.count
    }
    
    // Row display. Implementers should *always* try to reuse cells by setting each cell's reuseIdentifier and querying for available reusable cells with dequeueReusableCellWithIdentifier:
    // Cell gets various attributes set automatically based on table (separators) and data source (accessory views, editing controls)
    
    func tableView(tableView: UITableView, cellForRowAtIndexPath indexPath: NSIndexPath) -> UITableViewCell
    {
        let cell : UITableViewCell = tableView.dequeueReusableCellWithIdentifier("ProductCell")!
        let productRecord : Dictionary = productArray[indexPath.row] as Dictionary
        
        cell.textLabel?.text = productRecord["name"] as? String
        cell.detailTextLabel?.text = (productRecord["price"] as? String)! + (productRecord["date"] as? String)!
        cell.imageView?.image = UIImage(contentsOfFile: (productRecord["image"] as? String)!)
        
        
        return cell
    }
    
    func didFinishGeneratingDocument(path: String)
    {
        let documentInteractionController : UIDocumentInteractionController = UIDocumentInteractionController(URL: NSURL(fileURLWithPath: path))
        documentInteractionController.delegate = self
        documentInteractionController.presentPreviewAnimated(true)
        
        
    }
    
    func documentInteractionControllerViewControllerForPreview(controller: UIDocumentInteractionController) -> UIViewController {
        return self.navigationController!
    }
    

}

