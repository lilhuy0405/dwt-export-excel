const Excel = require('exceljs');

const moment = require("moment");

const express = require('express')
const cors = require("cors");
const fs = require("fs");
const app = express()
const port = 8888
const FormData = require('form-data');
const axios = require("axios");

function getKpiKey(obj) {
  if ('kpiKeys' in obj) {
    return {
      name: obj.kpiKeys[0] ? obj.kpiKeys[0].name : '',
      quantity: obj.kpiKeys[0] ? obj.kpiKeys[0].quantity : '',
      unit: obj.kpiKeys[0] ? obj.kpiKeys[0].unit.name : '',
    }
  }
  if ('kpi_keys' in obj) {
    return {
      name: obj.kpi_keys[0] ? obj.kpi_keys[0].name : '',
      quantity: obj.kpi_keys[0] ? obj.kpi_keys[0].pivot.quantity : '',
      unit: obj.kpi_keys[0] ? obj.kpi_keys[0].unit.name : '',
    }
  }
  return {
    name: '',
    quantity: '',
    unit: '',
  }
}

const fetch = require('node-fetch');
const querystring = require("querystring");

async function exportExcel(targetDetail, reportTask) {
  //this will be take from the api


  const workbook = new Excel.Workbook();
  const sample = await workbook.xlsx.readFile('sample.xlsx')
  const sheet = sample.getWorksheet(1);

  let startRow = 8;
  sheet.duplicateRow(startRow, targetDetail.length + reportTask.length - 1, true);
  await workbook.xlsx.writeFile('temp.xlsx');
  //write the data
  const tempFile = await workbook.xlsx.readFile('temp.xlsx')
  const toWriteSheet = tempFile.getWorksheet(1);
  //write the data
  let startWrite = 9;
  for (let i = 0; i < targetDetail.length + reportTask.length; i++) {
    const task = i < targetDetail.length ? targetDetail[i] : reportTask[i - targetDetail.length];
    const taskName = task.name;
    const {name: kpiKeyName, quantity: kpiKeyQuantity, unit: kpiKeyUnit} = getKpiKey(task);
    const executionPlan = task.executionPlan ?? '';
    const manday = 'manday' in task ? task.manday : task.manDay;
    const deadline = task.deadline ? moment(task.deadline).format('DD/MM/YYYY') : '';
    const employeeReportedQty = task.keysPassed;
    const calculatedKpi = task.kpiValue;
    const managerManday = task.managerManDay;
    const managerComment = task.managerComment ?? '';

    if (i === 0) {
      const row = toWriteSheet.getRow(startWrite - 1);
      row.getCell(1).value = i + 1;
      row.getCell(1).value = i + 1;
      row.getCell('B').value = taskName;
      row.getCell('G').value = kpiKeyName;
      row.getCell('L').value = manday;
      row.getCell('M').value = kpiKeyUnit;
      row.getCell('N').value = executionPlan;
      row.getCell('O').value = kpiKeyQuantity;
      row.getCell('P').value = deadline;
      row.getCell('Q').value = employeeReportedQty;
      row.getCell('R').value = calculatedKpi;
      row.getCell('T').value = managerManday;
      row.getCell('U').value = managerComment;

      continue;
    }
    const row = toWriteSheet.getRow(startWrite);
    toWriteSheet.mergeCells(`B${startWrite}:F${startWrite}`);
    toWriteSheet.mergeCells(`G${startWrite}:K${startWrite}`);

    row.getCell(1).value = i + 1;
    row.getCell(1).value = i + 1;
    row.getCell('B').value = taskName;
    row.getCell('G').value = kpiKeyName;
    row.getCell('L').value = manday;
    row.getCell('M').value = kpiKeyUnit;
    row.getCell('N').value = executionPlan;
    row.getCell('O').value = kpiKeyQuantity;
    row.getCell('P').value = deadline;
    row.getCell('Q').value = employeeReportedQty;
    row.getCell('R').value = calculatedKpi;
    row.getCell('T').value = managerManday;
    row.getCell('U').value = managerComment;

    startWrite++;
  }

  //save output
  await workbook.xlsx.writeFile('output.xlsx');
  return true;
}

app.use(cors());
app.use(express.json());
app.post('/get-excel', async (req, res) => {
  try {
    const {token, apiUrl} = req.body;
    console.log(token, apiUrl);
    const {
      reportTaskParams,
      targetDetailParams,
    } = req.body;
    const reportTaskQueryString = querystring.stringify(reportTaskParams);
    const targetDetailQueryString = querystring.stringify(targetDetailParams);

    const targetDetailRes = await fetch(`${apiUrl}/target-details?${targetDetailQueryString}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      }
    });
    if (!targetDetailRes.ok) {
      throw new Error('Get target detail failed');
    }
    const targetDetailData = await targetDetailRes.json();
    const targetDetail = targetDetailData.data.data;

    const reportTaskRes = await fetch(`${apiUrl}/report-tasks?${reportTaskQueryString}`, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      }
    });
    if(!reportTaskRes.ok) {
      throw new Error('Get report task failed');
    }

    const reportTaskData = await reportTaskRes.json();
    const reportTask = reportTaskData.data.data;


    // const {targetDetail, reportTask} = req.body;
    await exportExcel(targetDetail, reportTask);


    //send output file to hosst
    const uploadUrl = "https://report.sweetsica.com/api/report/upload";

    const formData = new FormData();
    formData.append('files', fs.createReadStream('output.xlsx'));
    const uploadRes = await fetch(uploadUrl, {
      method: 'POST',
      body: formData,
    });
    if (!uploadRes.ok) {
      const data = await res.text();
      console.log(data);
      throw new Error('Upload failed');
    }

    const data = await uploadRes.json();
    return res.status(200).send(data);

  } catch (err) {
    console.log(err);
    res.status(500).send({
      message: err.message,
    });
  }
})

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})
