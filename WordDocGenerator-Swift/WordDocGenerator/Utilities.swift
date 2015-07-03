//
//  Utilities.swift
//  WordDocGenerator-Swift
//
//  Created by AmitB on 02/07/15.
//  Copyright © 2015 Company_Name. All rights reserved.
//

import Foundation

 var documentsDirectory : NSString = ""

func getDocumentsDirectory()-> NSString
{
    
    if documentsDirectory.length <= 0
    {
        let paths : [AnyObject] = NSSearchPathForDirectoriesInDomains(NSSearchPathDirectory.DocumentDirectory, NSSearchPathDomainMask.UserDomainMask, true)
        documentsDirectory = paths[0] as! NSString
    }
    
    return documentsDirectory
//    static NSString *documentsDirectory= nil;
//    if(!documentsDirectory) {
//        //Search paths for directories in User domain
//        NSArray *paths = NSSearchPathForDirectoriesInDomains(NSDocumentDirectory,
//            NSUserDomainMask, YES);
//        documentsDirectory = paths[0];
//    }
//    return documentsDirectory;
}