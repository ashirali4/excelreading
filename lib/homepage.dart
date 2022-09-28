import 'dart:io';

import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:flutter_smart_dialog/flutter_smart_dialog.dart';

import 'loading.dart';

class HomePage extends StatefulWidget {
  const HomePage({Key? key}) : super(key: key);

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  File? file;
  List<int> units=[];
  List<String> processedIds = [];

  readExcelFile() async {
    if(file!=null){
      SmartDialog.showLoading(
        maskColor: Colors.black.withOpacity(.9),
        animationType: SmartAnimationType.scale,
        builder: (_) => CustomLoading(type: 1),
      );

     try{
       var bytes = file!.readAsBytesSync();
       Excel excel = Excel.decodeBytes(bytes);
       var myExcelFile = Excel.createExcel(); // automatically creates 1 empty sheet: Sheet1
       myExcelFile.rename('Sheet1', 'Worksheet');
       List<dynamic> excelRows=[];
       Sheet sheetObject = myExcelFile['Worksheet'];
       sheetObject.setColWidth(0, 20);
       sheetObject.setColWidth(1, 20);
       sheetObject.setColWidth(2, 50);
       sheetObject.setColWidth(3, 20);
       sheetObject.setColWidth(4, 20);
       sheetObject.setColWidth(5, 20);


       var rowsList  = excel.tables['Worksheet']!.rows;


       for (int excelFilE=1;excelFilE<rowsList.length ; excelFilE++) {
         String asin = rowsList[excelFilE][0]!.value;
         if(!processedIds.contains(asin)){
           int unitsdd=0;


           for(int a=0;a<rowsList.length ; a++){
             if(asin ==  rowsList[a][0]!.value){
               unitsdd =  unitsdd + int.parse(rowsList[a][3]!.value.toString());
               // if(excelFilE==a){
               //   print('Matching ASIN-- >' + asin);
               //
               // }else{
               //   print( '  -- Removed Index -- '+(a+1).toString());
               //  // exceltoWrite.removeRow('Worksheet', a+1);
               // }

             }
           }
           excelRows.add(rowsList[excelFilE]);
           units.add(unitsdd);
         }
         processedIds.add(asin);
         print("Lenght of " + processedIds.length.toString());
       }

       for(int a=0;a<excelRows.length;a++){
         processedIds.clear();
         String link =  'https://amazon.de/dp/'+excelRows[a][0].value;
         String asin =  excelRows[a][0].value;
         String itemDescription =  excelRows[a][2].value;
         //var units =  excelRows[a][3].value;
         var unitcost =  excelRows[a][4].value;
         var totalcost =  excelRows[a][5].value;
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("B"+(a+1).toString()), Formula.custom('HYPERLINK("$link","$asin")'),cellStyle: CellStyle(underline: Underline.Single,fontColorHex: 'FF3352FF') );
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("A"+(a+1).toString()), asin );
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("C"+(a+1).toString()), itemDescription );
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("D"+(a+1).toString()), units[a]);
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("E"+(a+1).toString()), unitcost );
         myExcelFile.updateCell('Worksheet', CellIndex.indexByString("F"+(a+1).toString()), totalcost );

       }

       var fileBytes = myExcelFile.save();

       File('C:\\Users\\ashir.muhammad\\Pictures\\excelFile\\result.xlsx')
         ..createSync(recursive: true)
         ..writeAsBytesSync(fileBytes!);


        // File(file!.path.toString())
        //  ..createSync(recursive: true)
        //  ..writeAsBytesSync(fileBytes!);


     }catch(e){
       print(e);
     }

      await Future.delayed(Duration(seconds: 2));
      SmartDialog.dismiss();
    }
  }


  uploadFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles();

    if (result != null) {
       file = File(result.files.single.path.toString());
      setState(() {
        pathFile = file!.path.toString();
      });
    } else {
      // User canceled the picker
    }
  }

  String pathFile = 'No File Selected';
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      // appBar: AppBar(
      //   title: Text('Upload your Excel File'),
      //   centerTitle: true,
      // ),
      body: Container(
        width: MediaQuery.of(context).size.width,
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          crossAxisAlignment: CrossAxisAlignment.center,
          children: [
            Container(
              height: 35,
              width: 150,
              child: TextButton(
                  child: Text(
                      "Upload Excel File".toUpperCase(),
                      style: TextStyle(fontSize: 14,color: Colors.white)
                  ),
                  style: ButtonStyle(
                      foregroundColor: MaterialStateProperty.all<Color>(Color(0xff004D77)),
                      backgroundColor: MaterialStateProperty.all<Color>(Color(0xff004D77)),
                      shape: MaterialStateProperty.all<RoundedRectangleBorder>(
                          RoundedRectangleBorder(
                              borderRadius: BorderRadius.zero,
                              side: BorderSide(color: Color(0xff004D77))
                          )
                      )
                  ),
                  onPressed: () {
                    uploadFile();
                    }
              ),
            ),
            SizedBox(height: 10,),
            Text(pathFile),
            SizedBox(height: 10,),
            Container(
              height: 35,
              width: 150,
              child: TextButton(
                  child: Text(
                      "Start Optimizing".toUpperCase(),
                      style: TextStyle(fontSize: 14,color: Colors.white)
                  ),
                  style: ButtonStyle(
                      foregroundColor: MaterialStateProperty.all<Color>(Color(0xff004D77)),
                      backgroundColor: MaterialStateProperty.all<Color>(Color(0xff004D77)),
                      shape: MaterialStateProperty.all<RoundedRectangleBorder>(
                          RoundedRectangleBorder(
                              borderRadius: BorderRadius.zero,
                              side: BorderSide(color: Color(0xff004D77))
                          )
                      )
                  ),
                  onPressed: () {
                   readExcelFile();
                  }
              ),
            ),
          ],
        ),
      ),
    );
  }
}
