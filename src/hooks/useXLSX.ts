
import { read } from 'xlsx'
import { ExcelFileValidType } from '../type/excel'
export default function useXLSX(){
    function getValidExcelTypeStr(){
        return Object.values(ExcelFileValidType).join("或")
    }
    function validExcelFileType(fileName:string){
        const isValid = Object.values(ExcelFileValidType).some(type=>{
            return fileName.includes(type)
        })
        if(!isValid) {
            alert(`請上傳excel專用的副檔名 : ${getValidExcelTypeStr()}`)
        }
        return isValid
    }
    async function handleExcelFileInput(event:React.ChangeEvent<HTMLInputElement>){
        if(!event.currentTarget.files?.length) return 
        if(!validExcelFileType(event.currentTarget.files[0].name)) return
        const file = event.currentTarget.files[0]
        const data = await file.arrayBuffer()
        return read(data)
    }
    return {
        getValidExcelTypeStr,
        handleExcelFileInput,
    }
}