const { dialog } = require('electron')
const fs = require('fs')
const path = require('path')
export const handleFileOpen = async (event, fileVal) => {
  console.log('filetype', fileVal)
  let properties = []
  if (fileVal === 'file') {
    properties = ['openFile']
  } else {
    properties = ['openDirectory']
  }
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties
  })
  if (!canceled) {
    if (fileVal === 'file') {
      return {
        filePath: filePaths[0],
        fileName: filePaths[0]
      }
    } else {
      const list = getDocxFiles(filePaths[0])
      // console.log('------all', list)

      return {
        filePath: list,
        fileName: filePaths[0]
      }
    }
  }
}
export const getDocxFiles = (dir) => {
  const docxFiles = []
  const allFileType = ['.docx', '.doc', '.xlsx', '.xls', '.pdf', '.txt', '.pptx']
  function walkDir(currentDir) {
    const files = fs.readdirSync(currentDir)
    files.forEach((file) => {
      const filePath = path.join(currentDir, file)
      const stat = fs.statSync(filePath)

      if (stat.isDirectory()) {
        // 如果是目录，递归调用
        walkDir(filePath)
      } else if (stat.isFile() && allFileType.includes(path.extname(file).toLowerCase())) {
        // 如果是 .docx 文件，添加到结果数组
        docxFiles.push(filePath)
      }
    })
  }
  walkDir(dir)
  return docxFiles
}
