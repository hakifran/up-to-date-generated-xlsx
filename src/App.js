import React, { Component } from 'react';
import logo from './logo.svg';
import './App.css';
import Workbook from 'react-excel-workbook'
import XLSX from 'xlsx';
import FileSaver from 'file-saver';
class App extends Component {

  constructor(props) {
    super(props);
    this.state = {
      reactexcelworkbookdata : {
        data1: [
          {
            foo: '123',
            bar: '456',
            baz: '789'
          },
          {
            foo: 'abc',
            bar: 'dfg',
            baz: 'hij'
          },
          {
            foo: 'aaa',
            bar: 'bbb',
            baz: 'ccc'
          },
          {
            foo: 'aaa',
            bar: 'bbb',
            baz: 'ccc'
          },
          {
            foo: 'aaa',
            bar: 'bbb',
            baz: 'ccc'
          }
      ],
      data2: [
        {
          aaa: 1,
          bbb: 2,
          ccc: 3
        },
        {
          aaa: 4,
          bbb: 5,
          ccc: 6
        }
      ]
    },

    xlsxFile: function (wbout) {
      var buf = new ArrayBuffer(wbout.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=wbout.length; ++i) view[i] = wbout.charCodeAt(i) & 0xFF;
      return buf;
    }

    }
  }

  render() {
    return (
      <div className="App">
        <ReactExcelWorkbook reactexcelworkbookdata={this.state.reactexcelworkbookdata}/>
        <br/>
        <br/>
        <XlsxArrayArrayToSheet xlsxFile={this.state.xlsxFile}/>
        <TableToSheet xlsxFile={this.state.xlsxFile}/>
      </div>
      );
  }
}

class ReactExcelWorkbook extends React.Component{
  constructor(props) {
    super(props)
  }
  render() {
    return(
      <Workbook filename="example_ReactExcelWorkbook.xlsx" element={<button className="btn btn-lg btn-primary">Try ReactExcelWorkbook!</button>}>
        <Workbook.Sheet data={this.props.reactexcelworkbookdata.data1} name="Sheet A">
          <Workbook.Column label="Foo" value="foo"/>
          <Workbook.Column label="Bar" value="bar"/>
        </Workbook.Sheet>
        <Workbook.Sheet data={this.props.reactexcelworkbookdata.data2} name="Another sheet">
          <Workbook.Column label="Double aaa" value={row => row.aaa * 2}/>
          <Workbook.Column label="Cubed ccc " value={row => Math.pow(row.ccc, 3)}/>
        </Workbook.Sheet>
      </Workbook>
    )
  }
}

class XlsxArrayArrayToSheet extends React.Component {
  constructor(props) {
    super(props);
    this.handleclick = this.handleclick.bind(this);;
  }


  handleclick() {
    var wb = XLSX.utils.book_new();
    wb.SheetNames.push("Test XlsxArrayArrayToSheet");
    var ws_data = [["hello", "world"]];
    var ws =  XLSX.utils.aoa_to_sheet(ws_data);
    wb.Sheets["Test XlsxArrayArrayToSheet"] = ws;
    var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
    var wbout = XLSX.write(wb,wopts);
    FileSaver.saveAs(new Blob([this.props.xlsxFile(wbout)],{type:""}), "test XlsxArrayArrayToSheet.xlsx")
  }

  render() {
    return(
      <div>
        <button onClick={this.handleclick}> Try XlsxArrayArrayToSheet</button>
      </div>
    );
  }
}

class TableToSheet extends React.Component {
  constructor(props) {
    super(props);
    this.handleclick = this.handleclick.bind(this);;
  }


  handleclick() {
    var wb = XLSX.utils.table_to_book(document.getElementById('mytable'), {sheet:"Table To Sheet"})
    var wbout =  XLSX.write(wb, {bookType: 'xlsx', bookSST:true, type: 'binary'});
    FileSaver.saveAs(new Blob([this.props.xlsxFile(wbout)], {type:""}), 'test TableToSheet.xlsx',)
  }

  render() {
    return(
      <div>
      <table id="mytable">
          <tbody>
            <tr>
              <td> 1 </td>
              <td> 2 </td>
              <td> 3 </td>
            </tr>
            <tr>
              <td> 4 </td>
              <td> 5 </td>
              <td> 6 </td>
            </tr>
          </tbody>
        </table>
        <button onClick={this.handleclick}>  Try TableToSheet </button>
      </div>
    );
  }
}

export default App;
