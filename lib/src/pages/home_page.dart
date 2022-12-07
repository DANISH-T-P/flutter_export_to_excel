import 'package:excel/excel.dart';
import 'package:flutter/material.dart';
import 'package:flutter_export_to_excel/src/pages/second_page.dart';


class HomePage extends StatefulWidget {
  const HomePage({Key? key}) : super(key: key);

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {

  _exportToExcel() {
    final excel = Excel.createExcel();
    final sheet = excel.sheets[excel.getDefaultSheet() as String];
    sheet!.setColWidth(1, 50);
    sheet.setColWidth(0, 5);
    sheet.setColAutoFit(2);
    sheet.setColAutoFit(3);


    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 0)).value = '0th column 0th row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 1)).value = '0th column first row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 2)).value = '0th column 2nd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 3)).value = '0th column 3rd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 1, columnIndex: 1)).value = 'First column first row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 1, columnIndex: 2)).value = 'First column 2nd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 1, columnIndex: 3)).value = 'First column 3rd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 2, columnIndex: 1)).value = '2nd column first row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 2, columnIndex: 2)).value = '2nd column 2nd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 2, columnIndex: 3)).value = '2nd column 3rd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 3, columnIndex: 1)).value = '3rd column first row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 3, columnIndex: 2)).value = '3rd column 2nd row';
    sheet.cell(CellIndex.indexByColumnRow(rowIndex: 3, columnIndex: 3)).value = '3rd column 3rd row';

    excel.save();
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Export to Excel'),
      ),
      body: Center(
        child: Container(
          height: 500,
          width: 500,
          decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(20),
              boxShadow: const [
                BoxShadow(
                  blurRadius: 10,
                  color: Colors.black,
                  offset: Offset(1,3),
                )
              ]
          ),
          child: Padding(
            padding: const EdgeInsets.only(top: 30,bottom: 20,right: 10,left: 10),
            child: Column(
              children: [
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Enter ID',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 200,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Name',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Class',
                        ),
                      ),
                    ),
                  ],
                ),
                const SizedBox(
                  height: 5,
                ),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Enter ID',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 200,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Name',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Class',
                        ),
                      ),
                    ),
                  ],
                ),
                const SizedBox(
                  height: 5,
                ),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Enter ID',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 200,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Name',
                        ),
                      ),
                    ),
                    Container(
                      // height: 100,
                      width: 100,
                      color: Colors.white10,
                      child: const TextField(
                        decoration: InputDecoration(
                          border: OutlineInputBorder(),
                          hintText: 'Class',
                        ),
                      ),
                    ),
                  ],
                ),
                const SizedBox(
                  height: 20,
                ),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      height: 50,
                      width: 200,
                      color: Colors.white10,
                      child: ElevatedButton(
                        onPressed: () {},
                        child: const Text('SAVE'),
                      ),
                    ),
                  ],
                ),
                const SizedBox(
                  height: 20,
                ),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      height: 50,
                      width: 200,
                      color: Colors.white10,
                      child: ElevatedButton(
                          onPressed: _exportToExcel, child: const Text('Export to EXCEL')),
                    ),
                  ],
                ),
                const SizedBox(
                  height: 20,
                ),
                Row(
                  mainAxisAlignment: MainAxisAlignment.spaceEvenly,
                  children: [
                    Container(
                      height: 50,
                      width: 200,
                      color: Colors.white10,
                      child: ElevatedButton(
                          onPressed: () {
                            Navigator.push(context, MaterialPageRoute(builder: (context) => SecondPage(),));
                          },
                          child: const Text('HomePage')),
                    ),
                  ],
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}