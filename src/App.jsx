import { useState } from 'react'
import './App.css'
import * as XLSX from 'xlsx'
import { AddVehicleBookingFromExcel } from './function/vehiclebooking';

function App() {
  const [excelFile, setExcelFile]=useState(null);
  const [excelFileError, setExcelFileError]=useState(null); 

  const [excelData, setExcelData]=useState(null);

  const fileType=['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
  const handleFile = (e)=>{
    let selectedFile = e.target.files[0];
    if(selectedFile){
      console.log(selectedFile.type);
      if(selectedFile&&fileType.includes(selectedFile.type)){
        let reader = new FileReader();
        reader.readAsArrayBuffer(selectedFile);
        reader.onload=(e)=>{
          setExcelFileError(null);
          setExcelFile(e.target.result);
        } 
      }
      else{
        setExcelFileError('อัพโหลดได้เเค่ไฟล์สกุล .xlsx');
        setExcelFile(null);
      }
    }
    else{
      console.log('plz select your file');
    }
  }

  const handleSubmit = async (e) => {
    e.preventDefault();
    if(excelFile!==null){
      const workbook = XLSX.read(excelFile,{type:'buffer'});
      const worksheetName = workbook.SheetNames[0];
      const worksheet=workbook.Sheets[worksheetName];

      // let checkOne = ''
      // let checkTwo = ''
      // let checkThree = ''
      // let checkFour = ''
      // let count = 3
      // while (checkOne != undefined && checkTwo != undefined && checkThree != undefined && checkFour != undefined) {
      //   checkOne = worksheet[`D${count + 1}`];
      //   checkTwo = worksheet[`E${count + 1}`];
      //   checkThree = worksheet[`I${count + 1}`];
      //   checkFour = worksheet[`L${count + 1}`];
      //   count = count + 1
      // }

      let allVehicleBooking = []
      
      for (let index = 2; index < 318; index++) {
        // ดึงข้อมูลออกมาจากไฟล์ทีละบรรทัด
        const id = worksheet[`A${index}`];
        const plateNumber = worksheet[`B${index}`];
        const client = worksheet[`D${index}`];
        const network = worksheet[`E${index}`];
        const team = worksheet[`F${index}`];
        const status = worksheet[`G${index}`];
        const remark = worksheet[`H${index}`];
        const issueDate = worksheet[`I${index}`];
        const problemIssue = worksheet[`J${index}`];

        // ถ้าข้อมูลมีการเว้นว่าง ให้เปลี่ยนเป็น null เพื่อป้องกัน Error
        let clientData
        let teamData  
        let issueDateData
        let problemIssueData
        let networkData
        let remarkData

        if (client == undefined) {
          clientData = 'N/A'
        } else {
          clientData = client.v
        }

        if (network == undefined) {
          networkData = 'N/A'
        } else {
          networkData = network.v
        }

        if (team == undefined) {
          teamData = 'N/A'
        } else {
          if (team.v == 'OPS LAS'){
            teamData = '1. ' + team.v
          } else if (team.v == 'OPS Forum'){
            teamData = '2. ' + 'OPS FORUM'
          } else if (team.v == 'OPS RETAIL'){
            teamData = '3. ' + team.v
          } else if (team.v == 'IMPLEMENT'){
            teamData = '4. ' + team.v
          } 
        }

        if (remark == undefined) {
          remarkData = null
        } else {
          remarkData = remark.v
        }

        if (issueDate == undefined) {
          issueDateData = null
        } else {
          issueDateData = issueDate.w
        }

        if (problemIssue == undefined) {
          problemIssueData = null
        } else {
          problemIssueData = problemIssue.v
        }

        console.log({
          plateNumber: plateNumber.v,
          client: clientData,
          network: networkData,
          team: teamData,
          status: status.v,
          remark: remarkData,
          issueDate: issueDateData,
          problemIssue: problemIssueData,
        });

        let vehicleBooking = ({
          id: id.v,
          plateNumber: plateNumber.v,
          client: clientData,
          network: networkData,
          team: teamData,
          status: status.v,
          remark: remarkData,
          issueDate: issueDateData,
          problemIssue: problemIssueData,
        })

        allVehicleBooking.push(vehicleBooking)
      }

      AddVehicleBookingFromExcel(allVehicleBooking)
    }
    else{
      setExcelData(null);
    }
  }

  console.log(excelFile);
  console.log(excelData);

  return (
    <div className="container p-5">
      {/* upload file section */}
      <div className='form'>
        <form className='form-group' autoComplete="off" onSubmit={handleSubmit}>
          <label><h5>Upload Excel file</h5></label>
          <br></br>
          <input type='file' className='form-control' onChange={handleFile} required></input>   
          {excelFileError&&<div className='text-danger' style={{marginTop:5+'px'}}>{excelFileError}</div>}     
          <button type='submit' className='btn btn-success' style={{marginTop:5+'px'}}>Submit</button>          
        </form>
      </div>
    </div>
  )
}

export default App
