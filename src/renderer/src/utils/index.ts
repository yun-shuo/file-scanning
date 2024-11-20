export interface KeywordResult {
  indices: number[]
  count: number
  examples: string[]
  filePath?: string
  fileName?: string
}
export function findKeywordOccurrences1(content: string, str: string): KeywordResult {
  const result: KeywordResult = { indices: [], count: 0, examples: [] }
  const keywords = str.split('|')
  let text = String(content)

  keywords.forEach((keyword) => {
    if (keyword.length === 0) return

    let index = text.indexOf(keyword)

    while (index !== -1) {
      result.indices.push(index)
      result.count++

      // 获取关键词的上下文示例
      if (result.examples.length < 10) {
        const contextLength = 20 // 上下文长度
        const start = Math.max(0, index - contextLength)
        const end = Math.min(text.length, index + keyword.length + contextLength)

        // 高亮关键词
        const beforeKeyword = text.substring(start, index)
        const afterKeyword = text.substring(index + keyword.length, end)
        const highlightedKeyword = `<span style="background-color: #ffdf00;">${keyword}</span>`
        const example = '...' + beforeKeyword + highlightedKeyword + afterKeyword + '...'
        result.examples.push(example)
      }

      index = text.indexOf(keyword, index + 1)
    }
  })

  // 去重并排序 indices
  result.indices = Array.from(new Set(result.indices)).sort((a, b) => a - b)

  // 如果 examples 数组中有重复项，去重
  result.examples = Array.from(new Set(result.examples))

  return result
}
export function findKeywordOccurrences(
  content: string,
  pattern: string,
  filePath: string
): KeywordResult | null {
  const result: KeywordResult = { indices: [], count: 0, examples: [] }

  // 解析模式，转换为正则表达式
  // const regexPattern = pattern
  //   .split('+') // 分割不同的组
  //   .map((group) => group.trim()) // 清除每个组的前后空格
  //   .map((group) => `(${group.replace(/\s+/g, '')})`) // 移除组内空格并加上括号
  //   .join('.*') // 组之间用 .* 连接，表示任何字符（包括换行符）出现任意次
  //   .replace(/\|/g, '|') // 确保 | 不被特殊处理
  const regexPattern = pattern
    .split('+') // 分割不同的组
    .map((group) => group.trim()) // 清除每个组的前后空格
    .map((group) => `(${group.replace(/\s+/g, '')})`) // 移除组内空格并加上括号
    .join('.*?')
  const regex = new RegExp(regexPattern, 'g')
  const testRegex = new RegExp(regexPattern)
  let fileName = filePath.split('\\').slice(-1)[0]
  // console.log(filePath, fileName)

  const highlightedFileName = fileName.replace(
    regex,
    (match) => `<span style="background-color: #ffdf00;">${match}</span>`
  )
  let match
  while ((match = regex.exec(content)) !== null) {
    const index = match.index
    const keyword = match[0]

    // 记录匹配的位置和增加计数
    result.indices.push(index)
    result.count++

    // 获取关键词的上下文示例
    // if (result.examples.length < 3) {
    const contextLength = 40 // 上下文长度
    const start = Math.max(0, index - contextLength)
    const end = Math.min(content.length, index + keyword.length + contextLength)

    // 高亮关键词
    const beforeKeyword = content.substring(start, index)
    const afterKeyword = content.substring(index + keyword.length, end)
    const highlightedKeyword = `<span style="background-color: #ffdf00;">${keyword}</span>`
    const example = '...' + beforeKeyword + highlightedKeyword + afterKeyword + '...'
    result.examples.push(example)
    // }

    // 防止无限循环
    if (match.index === regex.lastIndex) {
      regex.lastIndex++
    }
  }

  // 去重并排序 indices
  result.indices = Array.from(new Set(result.indices)).sort((a, b) => a - b)

  // 如果 examples 数组中有重复项，去重
  result.examples = Array.from(new Set(result.examples))
  result.fileName = highlightedFileName
  const isFilePathMatch = testRegex.test(filePath.split('\\').slice(-1)[0])
  result.filePath = filePath

  if (isFilePathMatch || result.count) {
    return result
  } else {
    return null
  }
}
export function findKeywordExcel1(jsonData: any[], keyword: string, filePath: string) {
  const list = [] as any[]
  const keywords = keyword.split('|')
  jsonData.forEach((content) => {
    let count = 0
    const examples: string[] = []

    content.data.forEach((row, rowIndex) => {
      for (const key in row) {
        // 遍历关键词数组
        for (const kw of keywords) {
          const cellValue = String(row[key])
          if (cellValue.includes(kw)) {
            if (examples.length < 10) {
              // 高亮关键词
              const highlightedExample = cellValue.replace(
                new RegExp(kw, 'g'),
                `<span style="background-color: #ffdf00;">$&</span>`
              )
              examples.push(highlightedExample)
            }
            count++
            // 如果已经找到一个匹配的关键词，则不再继续查找其他关键词
            break
          }
        }
      }
    })

    list.push({
      sheetName: content.sheetName,
      count,
      examples,
      filePath,
      fileName: filePath.split('\\').slice(-1)[0]
    })
  })

  return list
}
export function findPathHighlight(pattern: string, filePath: string) {
  const regexPattern = buildRegexPattern(pattern)
  const regex = new RegExp(regexPattern, 'g')
  const highlightedFilePath = filePath.replace(
    regex,
    (match) => `<span style="background-color: #ffdf00;">${match}</span>`
  )
  return highlightedFilePath
}
function buildRegexPattern(pattern: string): string {
  return pattern
    .split('+') // 分割不同的组
    .map((group) => group.trim()) // 清除每个组的前后空格
    .map((group) => `(${group.replace(/\s+/g, '')})`) // 移除组内空格并加上括号
    .join('.*') // 组之间用 .* 连接，表示任何字符（包括换行符）出现任意次
    .replace(/\|/g, '|') // 确保 | 不被特殊处理
}
export function findKeywordExcel(
  jsonData: any[],
  pattern: string,
  filePath: string,
  stats: any
): any[] {
  const list = [] as any[]
  const regexPattern = buildRegexPattern(pattern)
  const regex = new RegExp(regexPattern, 'g')
  const testRegex = new RegExp(regexPattern)
  jsonData.forEach((content) => {
    let count = 0
    const examples: string[] = []

    content.data.forEach((row, rowIndex) => {
      for (const key in row) {
        const cellValue = String(row[key])
        let match
        while ((match = regex.exec(cellValue)) !== null) {
          const keyword = match[0]
          // if (examples.length < 6) {
          // 截取匹配句子，确保不超过20个字符
          const truncatedSentence = cellValue.slice(match.index, match.index + 40)
          // 高亮关键词
          const highlightedExample = truncatedSentence.replace(
            new RegExp(keyword, 'g'),
            `<span style="background-color: #ffdf00;">$&</span>`
          )
          examples.push(highlightedExample)
          // }
          count++
          // 防止无限循环
          if (match.index === regex.lastIndex) {
            regex.lastIndex++
          }
        }
      }
    })
    const fileName = filePath.split('\\').slice(-1)[0]
    const isFilePathMatch = testRegex.test(fileName)
    const highlightedFileName = fileName.replace(
      regex,
      (match) => `<span style="background-color: #ffdf00;">${match}</span>`
    )
    if (count || isFilePathMatch) {
      list.push({
        sheetName: content.sheetName,
        count,
        examples,
        filePath: filePath,
        fileSize: stats.size,
        fileName: highlightedFileName,
        mtime: stats.mtime
      })
    }
  })

  return list
}
export const removeHtmlTags = (str) => str.replace(/<[^>]+>/g, '')
export const generateReport = (obj) => {
  const wordResults = obj.data.wordTableData
    .filter((item) => item.count)
    .map(
      (item) => `
      - 出现在文件“${item.filePath}”中，共出现了${item.count}次。<br/>`
    )
    .join('\n')

  const excelResults = obj.data.excelTableData
    .filter((item) => item.count)
    .map(
      (item) => `
      - 在文件“${item.filePath}”的子表“${item.sheetName}”中，共出现了${item.count}次。<br/>`
    )
    .join('\n')

  const pdfResults = obj.data.pdfTableData
    .filter((item) => item.count)
    .map(
      (item) => `
      - 在文件“${item.filePath}”中，共出现了${item.count}次。<br/>`
    )
    .join('\n')

  const txtResults = obj.data.txtTableData
    .filter((item) => item.count)
    .map(
      (item) => `
      - 在文件“${item.filePath}”中，共出现了${item.count}次。<br/>`
    )
    .join('\n')

  return `
根据本次扫描，我们对指定文件夹内的Word文档、Excel文件、PDF文件和TXT文件进行了关键字检测。扫描结果显示：
<br/>
<br/>
Word文档中：
${wordResults || '暂无'}
<br/>
<br/>
Excel文件中：
${excelResults || '暂无'}
<br/>
<br/>
PDF文件中：
${pdfResults || '暂无'}
<br/>
<br/>
TXT文件中：
${txtResults || '暂无'}
<br/>
<br/>
总体来看，建议对这些文件进行进一步审查，以确保内容的准确性和合规性。`
}
export function formatFileSize(bytes) {
  if (bytes === 0) return '0 B'

  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))

  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}
