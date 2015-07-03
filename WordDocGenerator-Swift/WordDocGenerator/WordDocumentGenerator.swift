//
//  WordDocumentGenerator.swift
//  WordDocGenerator-Swift
//
//  Created by AmitB on 02/07/15.
//  Copyright © 2015 Company_Name. All rights reserved.
//

import Foundation
import UIKit

protocol WordDocumentGeneratorDelegate : NSObjectProtocol
{
    func didFinishGeneratingDocument(path : String)
    
}

class WordDocumentGenerator: NSObject {
    private let XML_RELATIONSHIP : String = "<Relationship Id=\"%@\" Type=\"%@\" Target=\"%@\"/>"
    
    private var docXML : String?
    private var imagePaths : [String]?
    private var destinationDirectory : String?
    
    
    static let sharedInstance: WordDocumentGenerator =
    {
        return WordDocumentGenerator()
    }()
    
    func generateWordDocumentWithXML(docXML:String, imagePaths:[String], destinationDirectory:String, delegate:WordDocumentGeneratorDelegate)
    {
        self.docXML = docXML
        self.imagePaths = imagePaths
        self.destinationDirectory = destinationDirectory
        
        //Step 1
        //Copy Contents from bundle to destination directory.
        self.copyDirectory(destinationDirectory)
        
        //Step 2
        //Unzip word contents.
        self.unzipWordContents(getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory as String))
        
        //Step 3
        //Create document xml
        self.createDocumentXML(docXML, imagePaths: imagePaths, destinationDirectory:getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory))
        
        //Step 4
        //Zip word contents
        self.zipWordContents(getDocumentsDirectory().stringByAppendingPathComponent (destinationDirectory as String))
        
        //Step 5
        //Move zip file to .docx file.
        do {
            
                try NSFileManager.defaultManager().copyItemAtPath(getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory as String).stringByAppendingPathComponent("ProductList.zip"), toPath: getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory as String).stringByAppendingPathComponent("ProductList.docx"))

        }catch let error as NSError
        {
            print("error cathced is \(error)")
        }
        if delegate.respondsToSelector("didFinishGeneratingDocument:")
        {
            delegate.didFinishGeneratingDocument(getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory as String) .stringByAppendingPathComponent("ProductList.docx"))
        }
        
    }
    
    func copyDirectory (destinationDirectory : NSString)
    {
        let fileManager : NSFileManager = NSFileManager.defaultManager()
        
        let destinationPath : String = getDocumentsDirectory().stringByAppendingPathComponent(destinationDirectory as String)
        
        if fileManager.fileExistsAtPath(destinationPath) == false
        {
            do {
            try fileManager.createDirectoryAtPath(destinationPath, withIntermediateDirectories: true, attributes: nil)
            }catch let error as NSError
            {
                print("error in copying directory \(error)")
            }
        }else
        {
            print("Directory exists \(destinationPath)")
        }
        
        let sourcePath : String? = (NSBundle.mainBundle().resourcePath!.stringByAppendingPathComponent("WordContents.zip"))
        if fileManager.fileExistsAtPath(sourcePath!)
        {
            do{
                try fileManager.copyItemAtPath(sourcePath!, toPath: destinationPath.stringByAppendingPathComponent("WordContents.zip"))
            }catch let error as NSError
            {
                print("Error while coping wordcontents \(error)")
            }
        }
        
    }
    
    func unzipWordContents (destinationDirectory : String)
    {
        let zipFilePath : String = destinationDirectory.stringByAppendingPathComponent("WordContents.zip")
        //unzip file at above path
        SSZipArchive.unzipFileAtPath(zipFilePath, toDestination: destinationDirectory)
        
        if NSFileManager.defaultManager().fileExistsAtPath(zipFilePath)
        {
            do {
                try NSFileManager.defaultManager().removeItemAtPath(zipFilePath)
            }catch let error as NSError
            {
                print("Error while unzipping contents \(error)")
            }
        }
    }
    
    
    func zipWordContents (destinationDirectory:String)
    {
        SSZipArchive.createZipFileAtPath(NSTemporaryDirectory().stringByAppendingPathComponent("ProductList.zip"), withContentsOfDirectory: destinationDirectory)
        do {
            
            
            try NSFileManager.defaultManager().copyItemAtPath(NSTemporaryDirectory().stringByAppendingPathComponent("ProductList.zip"), toPath: destinationDirectory.stringByAppendingPathComponent("ProductList.zip"))
        }catch let error as NSError
        {
            print("Error while zipping word content \(error)")
        }
    }
    
    func createDocumentXML (docXML : NSString , imagePaths : [String], destinationDirectory:String)
    {
        let mediaDirectory : String = (destinationDirectory.stringByAppendingPathComponent("/word/media"))
        var imageNumber : Int = 1
        for imagePath in imagePaths {
            let newPath : String = String(format: "%@/image%i.jpg", mediaDirectory, imageNumber)
            let imageView : UIImageView = UIImageView(frame: CGRectMake(0, 0, 753, 1004))
            imageView.image = UIImage(contentsOfFile: imagePath as String)
            UIGraphicsBeginImageContext(CGSizeMake(753, 1004))
            imageView.layer.renderInContext(UIGraphicsGetCurrentContext())
            let viewImage : UIImage = UIGraphicsGetImageFromCurrentImageContext()
            UIGraphicsEndImageContext()
            
            UIImageJPEGRepresentation(viewImage, 1.0)?.writeToFile(newPath, atomically: true)
            
            imageNumber++
            
        }
        
        var imageIDs: [String] = self.createRelationshipDocument(imagePaths, destinationDirectory:destinationDirectory)
        var imageCounter : Int = 0;
        var documentXML = docXML
        for imagePath in imagePaths
        {
            documentXML = documentXML.stringByReplacingOccurrencesOfString(imagePath as String, withString: imageIDs[imageCounter])
            imageCounter++
        }
        
        if documentXML.length > 0
        {
            let docXMLPath : String = (destinationDirectory.stringByAppendingPathComponent("word").stringByAppendingPathComponent("document.xml"))
            do {
                
                try documentXML .writeToFile(docXMLPath, atomically: true, encoding: NSUTF8StringEncoding)
            }catch let error as NSError
            {
                print("Error while writing document XML to file \(error)")
            }
            
        }
        
        
    }
    
    func createRelationshipDocument(imagePathList:[String], destinationDirectory:String)->[String]
    {
        var imageIDs : [String] = [String]()
        
        var relXML : String = String("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">")
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml" )
        
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId2",
            "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects", "stylesWithEffects.xml")
        
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId3", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml")
        
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId4",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings", "webSettings.xml")
        
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId7",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "fontTable.xml")
        
        relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, "rId8",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml")
        var imageIdCounter : Int = 9
        
        for(var i = 0; i < imagePathList.count; i++)
        {
            let imageName : String = String(format: "media/image%d.jpg",i+1)
            let imageId : String = String(format: "rId%d", imageIdCounter)
            relXML = relXML.stringByAppendingFormat(XML_RELATIONSHIP, imageId,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", imageName)
            imageIDs.append(imageId)
            imageIdCounter++
        }
        
        relXML = relXML.stringByAppendingString("</Relationships>")
        
        let relsFilePath : String = destinationDirectory.stringByAppendingString("/word/_rels/document.xml.rels")
        do{
            try relXML.writeToFile(relsFilePath, atomically: true, encoding: NSUTF8StringEncoding)
        }catch let error as NSError
        {
            print("Error while creating relationship of style, setting, fonts, table of XML \(error)")
        }
        return imageIDs;
    }
    
}
