import mammoth from 'mammoth'
const fs = require('fs')
// const mammoth = require('mammoth')
// const XLSX = require('xlsx')
import XLSX from 'xlsx'
import pdf from 'pdf-parse'
const Extractor = require('word-extractor')
const officeParser = require('officeparser')
// const pdf = require('pdf-parse')
/**
 * 读取 Office 文件内容
 * @param {string} filePath - 文件路径
 * @returns {Promise<string|Array>} - 返回文件内容（Word）或数据（Excel）
 */
const extractor = new Extractor()
export async function readOfficeFile(filePath) {
  return new Promise((resolve, reject) => {
    const fileExtension = filePath.split('.').pop().toLowerCase()
    const stats = fs.statSync(filePath)
    if (fileExtension === 'docx' || fileExtension === 'doc') {
      // 处理 Word 文件
      // fs.readFile(filePath, (err, data) => {
      //   // console.log(err, data)
      //   if (err) {
      //     return resolve(null)
      //   }
      //   mammoth
      //     .extractRawText({ buffer: data })
      //     .then((result) => resolve({ fileContent: result.value, stats })) // 返回文本内容
      //     .catch(reject)
      // })
      const extracted = extractor.extract(filePath)

      extracted
        .then(function (doc) {
          resolve({ fileContent: doc.getBody(), stats })
        })
        .catch(reject)
    } else if (fileExtension === 'txt') {
      fs.readFile(filePath, 'utf8', (err, data) => {
        if (err) {
          console.error('读取文件时出错:', err)
          return resolve(null)
        }
        resolve({ fileContent: data, stats })
      })
    } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      // 处理 Excel 文件
      try {
        const workbook = XLSX.readFile(filePath)
        const sheetNames = workbook.SheetNames
        const data = []

        sheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName]
          const jsonData = XLSX.utils.sheet_to_json(worksheet)
          data.push({ sheetName, data: jsonData })
        })

        resolve({ fileContent: data, stats })
      } catch (error) {
        console.error('Error reading Excel file:', error)
        return resolve(null)
      }
    } else if (fileExtension === 'pdf') {
      // 处理 Pdf 文件
      officeParser
        .parseOfficeAsync(filePath)
        .then((data) => resolve({ fileContent: data, stats }))
        .catch((err) => {
          console.error(err)
          return resolve(null)
        })
      // fs.readFile(filePath, (err, data) => {
      //   if (err) {
      //     console.error('Error reading the file:', err)
      //     return resolve(null)
      //   }
      //   // 解析PDF文件
      //   pdf(data)
      //     .then(function (data) {
      //       // data.text 是PDF文件中的文本内容
      //       resolve({ fileContent: data.text, stats })
      //     })
      //     .catch(function (error) {
      //       console.error('Error parsing the PDF:', error)
      //       return resolve(null)
      //     })
      // })

      // resolve(data) // 返回每个工作表的数据
    } else if (fileExtension === 'pptx') {
      // 处理 Pdf 文件
      officeParser
        .parseOfficeAsync(filePath)
        .then((data) => resolve({ fileContent: data, stats }))
        .catch((err) => {
          console.error(err)
          return resolve(null)
        })
    } else {
      resolve({ fileContent: null, stats: null })
      // reject(new Error('Unsupported file type'))
    }
  })
}

// 使用示例
// readOfficeFile('path/to/your/file.docx')
//     .then(content => console.log(content))
//     .catch(err => console.error(err));

// readOfficeFile('path/to/your/file.xlsx')
//     .then(data => console.log(data))
//     .catch(err => console.error(err));
