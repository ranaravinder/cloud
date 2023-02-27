var fs = require('fs');
const Excel = require('exceljs')

const { type } = require('os');
var obj = JSON.parse(fs.readFileSync('adf-pipeline.json', 'utf8'));
const pipelines=obj.value
var data=[]

pipelines.forEach((el)=>
{
  let {name,properties} =el
  let {description,activities}=properties
  let activity,activityType,activitydependsOn,linkedService,input,output

  if (activities){
    activities.forEach((el)=>{
        let {dependsOn,linkedServiceName}=el     
        activitydependsOn=""
        activity=el.name
        activityType=el.type
        input=el.inputs
        output=el.outputs
         if (linkedServiceName){
            linkedService=linkedServiceName.referenceName       
         }
    
         if(dependsOn){
            dependsOn.forEach((e)=>{
                activitydependsOn= e.activity
            })     
        }
        
        let pipe={
            pipeline: name,
            description:description===undefined?"":description,
            activity: activity,
            activityType: activityType,
            dependsOn:activitydependsOn,
            linkedService:linkedService,
            input:input ===undefined?"":input[0].referenceName,
            output:output ===undefined?"":output[0].referenceName      
          }
          data.push(pipe)        
      })
  } 
   else{
    let pipe={
        pipeline: name,
        description:description===undefined?"":description,
        activity: activity,
        activityType: activityType,
        dependsOn:activitydependsOn,
        linkedService:linkedService,
        input:input,
        output:output    
      }
      data.push(pipe)

   }
}
)
console.clear()
console.dir(data)
let workbook = new Excel.Workbook()
let worksheet = workbook.addWorksheet('pipelines')
worksheet.columns = [
    {header: 'Pipeline Name', key: 'pipeline'},
    {header: 'Description', key: 'description'},
    {header: 'Activity', key: 'activity'},
    {header: 'Activity Type', key: 'activityType'},
    {header: 'Depended Activity', key: 'dependsOn'},
    {header: 'Linked Service', key: 'linkedService'},
    {header: 'Inputs', key: 'input'},
    {header: 'Outputs', key: 'output'}
  ]

data.forEach((e, index) => {
    // row 1 is the header.
    const rowIndex = index + 2
    worksheet.addRow({
      ...e     
    })
  })
  workbook.xlsx.writeFile('adf_pipelines.xlsx')