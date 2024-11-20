<script setup lang="ts">
import { computed, onMounted, ref, watch } from 'vue'
import dayjs from 'dayjs'
import {
  findKeywordExcel,
  findKeywordOccurrences,
  KeywordResult,
  generateReport,
  formatFileSize,
  removeHtmlTags
} from './utils/index'
import { useToast } from 'primevue/usetoast'
const toast = useToast()
const ipcHandle = () => window.electron.ipcRenderer.send('ping')
const allType = ['docx', 'doc', 'xlsx', 'xls', 'pdf', 'txt', 'pptx']
function readOfficeFile(filePath) {
  return new Promise(async (resolve, reject) => {
    // console.log('当前路径', filePath)
    const fileExtension = filePath.split('.').pop().toLowerCase()
    // console.log('文件类型', fileExtension)

    if (!allType.includes(fileExtension)) {
      resolve(true)
      return
    }
    const data = await window.electron.ipcRenderer.invoke('read-office-file', filePath)
    // console.log(data)

    if (!data) {
      resolve(true)
      return
    }
    const { fileContent, stats } = data
    // console.log(stats)
    //#region 解析word
    if (fileExtension === 'docx' || fileExtension === 'doc') {
      // console.log(filePath)
      const result = findKeywordOccurrences(fileContent, keyword.value, filePath)
      // console.log(content)
      result &&
        wordTableData.value.push({
          ...result,
          fileSize: stats.size,
          mtime: stats.mtime
        })
      resolve(true)
    }
    //#endregion
    //#region 解析word
    if (fileExtension === 'xlsx' || fileExtension === 'xls') {
      const result = findKeywordExcel(fileContent, keyword.value, filePath, stats)
      excelTableData.value.push(...result)
      resolve(true)
    }
    //#endregion
    //#region 解析word
    if (fileExtension === 'pdf') {
      const result = findKeywordOccurrences(fileContent, keyword.value, filePath)
      result &&
        pdfTableData.value.push({
          ...result,
          fileSize: stats.size,
          mtime: stats.mtime
        })
      resolve(true)
    }
    //#endregion
    //#region 解析txt
    if (fileExtension === 'txt') {
      const result = findKeywordOccurrences(fileContent, keyword.value, filePath)
      // console.log('txt结果', result)

      result &&
        txtTableData.value.push({
          ...result,
          fileSize: stats.size,
          mtime: stats.mtime
        })
      resolve(true)
    }
    //#endregion
    //#region 解析txt
    if (fileExtension === 'pptx') {
      const result = findKeywordOccurrences(fileContent, keyword.value, filePath)
      result &&
        pptTableData.value.push({
          ...result,
          fileSize: stats.size,
          mtime: stats.mtime
        })
      resolve(true)
    }
    //#endregion
  })
}
const input = ref('')
const keyword = ref('')
const wordTableData = ref([] as any[])
const excelTableData = ref([] as any[])
const pdfTableData = ref([] as any[])
const txtTableData = ref([] as any[])
const pptTableData = ref([] as any[])
const fileNum = ref(0)
const doneFileNum = ref(0)
const currentFileName = ref('')
const scanTimeStart = ref(new Date().getTime())
const scanTimeEnd = ref(new Date().getTime())
const directoryName = ref('')

const submit = async () => {
  const { filePath: resFile, fileName } = await window.electron.ipcRenderer.invoke(
    'openFile',
    fileVal.value
  )
  // console.log(resFile)
  toast.add({
    severity: 'info',
    summary: '文件读取中',
    detail: `请耐心等待`,
    life: 2000
  })
  directoryName.value = fileName
  // input.value = filePath
  wordTableData.value = []
  excelTableData.value = []
  pdfTableData.value = []
  txtTableData.value = []
  pptTableData.value = []
  fileNum.value = 0
  doneFileNum.value = 0
  scanTimeStart.value = new Date().getTime()
  if (Array.isArray(resFile)) {
    fileNum.value = resFile.length
    for (let index = 0; index < resFile.length; index++) {
      const path = resFile[index]
      doneFileNum.value++
      // console.log(doneFileNum.value)
      currentFileName.value = path
      await readOfficeFile(path)
    }
  } else {
    fileNum.value = 1
    doneFileNum.value++
    await readOfficeFile(resFile)
  }
  scanTimeEnd.value = new Date().getTime()
}
const fileVal = ref('directory')
const showFileType = ref('docx')
watch(
  () => showFileType.value,
  (newVal) => {
    if (!newVal) {
      showFileType.value = 'docx'
    }
  }
)
const compSelectOptions = computed(() => {
  const list = [
    { name: 'word', value: 'docx' },
    { name: 'excel', value: 'xlsx' },
    { name: 'pdf', value: 'pdf' },
    { name: 'txt', value: 'txt' },
    { name: 'ppt', value: 'ppt' }
  ] as any[]
  wordTableData.value.length && (list[0].num = wordTableData.value.length)
  excelTableData.value.length && (list[1].num = excelTableData.value.length)
  pdfTableData.value.length && (list[2].num = pdfTableData.value.length)
  txtTableData.value.length && (list[3].num = txtTableData.value.length)
  pptTableData.value.length && (list[4].num = pptTableData.value.length)
  return list
})
const openLink = (url: string) => {
  window.electron.ipcRenderer.invoke('open-link', url.replace(/<[^>]+>/g, ''))
}
const addScheme = async () => {
  const obj = {
    keyword: keyword.value,
    directoryName: directoryName.value,
    data: {
      wordTableData: wordTableData.value,
      excelTableData: excelTableData.value,
      pdfTableData: pdfTableData.value,
      txtTableData: txtTableData.value,
      pptTableData: pptTableData.value
    }
  }
  const num = await window.electron.ipcRenderer.invoke('store-scheme-add', JSON.stringify(obj))
  toast.add({
    severity: 'success',
    summary: '保存成功',
    detail: `已保存“${keyword.value}”方案`,
    life: 3000
  })
}
const schemeList = ref([] as any[])
const getScheme = async () => {
  schemeList.value = await window.electron.ipcRenderer.invoke('store-scheme-get')
  console.log('全部方案', schemeList.value)
}
onMounted(() => {
  // getScheme()
})

const visibleList = ref(false)
const visibleTop = ref(false)
const moreDataList = ref([])
const moreDataTitle = ref('')
const showMoreData = (row: any) => {
  visibleTop.value = true
  moreDataList.value = row.examples
  moreDataTitle.value = row.filePath
}
const openList = async () => {
  visibleList.value = true
  getScheme()
}
const lookDetail = (obj: any) => {
  const { keyword: str, data } = obj
  keyword.value = str
  visibleList.value = false
  wordTableData.value = data.wordTableData
  excelTableData.value = data.excelTableData
  pdfTableData.value = data.pdfTableData
  txtTableData.value = data.txtTableData
  pptTableData.value = data.pptTableData || []
}
const msgData = ref('')
const getMsg = (obj: any) => {
  msgData.value = generateReport(obj)
  visibleDialog.value = true
  titleDialog.value = obj.keyword + '-方案扫描报告'
}
const visibleDialog = ref(false)
const titleDialog = ref('')
</script>

<template>
  <div class="p-x-24px p-y-18px w-100%">
    <Toast />
    <Dialog v-model:visible="visibleDialog" modal :header="titleDialog" :style="{ width: '80%' }">
      <div v-html="msgData"></div>
    </Dialog>
    <Drawer v-model:visible="visibleList" style="width: 80vw" header="方案列表" position="right">
      <DataTable :value="schemeList" showGridlines scrollable stripedRows scrollHeight="80vh">
        <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.index + 1 }}
          </template></Column
        >
        <Column field="keyword" header="扫描文件/目录">
          <template #body="slotProps"> {{ slotProps.data.directoryName }} </template>
        </Column>
        <Column field="keyword" header="方案名称">
          <template #body="slotProps"> {{ slotProps.data.keyword }} </template>
        </Column>
        <Column field="wordTableData" header="word匹配结果" style="width: 100px">
          <template #body="slotProps"> {{ slotProps.data.data.wordTableData.length }}个 </template>
        </Column>
        <Column field="excelTableData" header="excel匹配结果" style="width: 100px">
          <template #body="slotProps"> {{ slotProps.data.data.excelTableData.length }}个 </template>
        </Column>
        <Column field="pdfTableData" header="pdf匹配结果" style="width: 100px">
          <template #body="slotProps"> {{ slotProps.data.data.pdfTableData.length }}个 </template>
        </Column>
        <Column field="txtTableData" header="txt匹配结果" style="width: 100px">
          <template #body="slotProps"> {{ slotProps.data.data.txtTableData.length }}个 </template>
        </Column>
        <Column field="pptTableData" header="ppt匹配结果" style="width: 100px">
          <template #body="slotProps">
            {{ slotProps.data.data.pptTableData?.length || 0 }}个
          </template>
        </Column>
        <Column field="fileName" header="操作" style="width: 120px">
          <template #body="slotProps">
            <div class="flex">
              <div class="color-blue cursor-pointer" @click="lookDetail(slotProps.data)">查看</div>
              <div class="color-blue cursor-pointer m-l-12px" @click="getMsg(slotProps.data)">
                报告
              </div>
            </div>
          </template>
        </Column>
      </DataTable>
    </Drawer>
    <Drawer
      v-model:visible="visibleTop"
      :header="moreDataTitle"
      position="top"
      style="height: auto"
    >
      <div class="max-h-70vh overflow-y-auto">
        <p class="m-y-12px color-#333" v-for="item in moreDataList" v-html="item"></p>
      </div>
    </Drawer>
    <Card>
      <!-- <template #title>文件扫描</template> -->
      <template #content>
        <div class="flex">
          <div class="w-full flex flex-1 items-center">
            <Select
              v-model="fileVal"
              :options="[
                {
                  name: '文件夹',
                  value: 'directory'
                },
                {
                  name: '文件',
                  value: 'file'
                }
              ]"
              optionLabel="name"
              option-value="value"
              size="small"
              class="m-r-12px"
            />
            <InputText
              size="small"
              type="text"
              placeholder="可用形如“(关键词一|关键词二)+(关键词三|关键词四)”的形式进行搜索"
              v-model="keyword"
              class="flex-1"
            />
          </div>
          <div class="m-l-24px h-40px flex justify-end">
            <Button
              :label="fileVal == 'directory' ? '选择扫描文件夹' : '选择扫描文件'"
              class="w-180px"
              @click="submit"
              :disabled="!keyword"
            />
            <Button
              label="保存方案"
              class="w-120px m-l-8px"
              :disabled="!keyword"
              @click="addScheme"
            />
            <Button icon="pi pi-list" class="m-l-12px" @click="openList" />
            <!-- <Button label="获取方案" class="w-120px m-l-8px" @click="getScheme" /> -->
          </div>
        </div>
      </template>
    </Card>

    <div class="m-y-12px">
      <div class="flex justify-between items-center">
        <SelectButton
          v-model="showFileType"
          optionLabel="name"
          optionValue="value"
          :options="compSelectOptions"
        >
          <template #option="slotProps">
            <OverlayBadge
              v-if="slotProps.option.num"
              :value="slotProps.option.num"
              size="small"
              severity="warn"
            >
              <span>{{ slotProps.option.name }}</span>
            </OverlayBadge>
          </template>
        </SelectButton>
        <div class="flex items-center" v-if="fileNum && doneFileNum != fileNum">
          <span class="max-w-50vw line-clamp-1 h-24px m-r-12px color-dark">
            {{ currentFileName }}
          </span>
          <div class="bg-#fff w-30vw rounded-16px h-18px">
            <div
              class="bg-green rounded-16px h-100%"
              :style="`width:${(doneFileNum / fileNum) * 100}% ;`"
            ></div>
          </div>
          <span class="m-l-12px color-dark"> {{ doneFileNum }}/{{ fileNum }} </span>

          <!-- <ProgressBar class="flex-1 m-x-120px" :value="(doneFileNum / fileNum) * 100">
            {{ doneFileNum }}/{{ fileNum }}
          </ProgressBar> -->
        </div>

        <Message class="" v-if="fileNum && scanTimeEnd - scanTimeStart > 0">
          <span>扫描文件数：{{ fileNum }}个;&emsp;</span>
          <span>扫描耗时：{{ ((scanTimeEnd - scanTimeStart) / 1000).toFixed(2) }}秒</span>
        </Message>
      </div>
    </div>
    <div class="m-t-12px" v-if="showFileType === 'txt'">
      <DataTable
        v-if="txtTableData.length"
        :value="txtTableData"
        tableStyle="min-width: 50%"
        showGridlines
        scrollable
        stripedRows
        selectionMode="single"
        scrollHeight="calc(100vh - 200px)"
      >
        <!-- <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.index + 1 }}
          </template></Column
        > -->
        <Column field="fileName" header="文件名" style="font-weight: 500; width: 100px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              contenteditable
              v-html="slotProps.data.fileName"
            ></div> </template
        ></Column>
        <Column field="fileName" header="文件路径" style="width: 200px">
          <template #body="slotProps">
            <div
              class="m-b-12px color-blue cursor-pointer"
              @click="openLink(slotProps.data.filePath)"
              v-html="slotProps.data.filePath"
            ></div>
          </template>
        </Column>
        <Column field="fileSize" header="文件大小" style="width: 100px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">{{ formatFileSize(slotProps.data.fileSize) }}</div>
          </template>
        </Column>
        <Column field="fileSize" header="修改时间" style="width: 180px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">
              {{ dayjs(slotProps.data.mtime).format('YYYY-MM-DD HH:mm:ss') }}
            </div>
          </template>
        </Column>
        <Column
          field="count"
          header="出现次数"
          style="width: 120px; font-weight: bold"
          sortable
        ></Column>
        <Column field="examples" header="示例">
          <template #body="slotProps">
            <div>
              <div
                class="m-b-12px"
                contenteditable
                v-for="item in slotProps.data.examples.slice(0, 2)"
                v-html="item"
              ></div>
              <div
                v-if="slotProps.data.examples.length > 2"
                class="color-blue cursor-pointer text-center text-12px"
                @click="showMoreData(slotProps.data)"
              >
                查看全部
              </div>
            </div>
          </template></Column
        >
      </DataTable>
      <div
        class="flex items-center justify-center"
        v-if="!txtTableData.length"
        style="height: calc(100vh - 200px)"
      >
        <img src="./assets/images/empty1.png" alt="" />
      </div>
    </div>
    <div class="m-t-12px" v-if="showFileType === 'ppt'">
      <DataTable
        v-if="pptTableData.length"
        :value="pptTableData"
        tableStyle="min-width: 50%"
        showGridlines
        scrollable
        stripedRows
        selectionMode="single"
        scrollHeight="calc(100vh - 200px)"
      >
        <!-- <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.index + 1 }}
          </template></Column
        > -->
        <Column field="fileName" header="文件名" style="font-weight: 500; width: 100px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              contenteditable
              v-html="slotProps.data.fileName"
            ></div> </template
        ></Column>
        <Column field="fileName" header="文件路径" style="width: 200px">
          <template #body="slotProps">
            <div
              class="m-b-12px color-blue cursor-pointer"
              @click="openLink(slotProps.data.filePath)"
              v-html="slotProps.data.filePath"
            ></div>
          </template>
        </Column>
        <Column field="fileSize" header="文件大小" style="width: 100px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">{{ formatFileSize(slotProps.data.fileSize) }}</div>
          </template>
        </Column>
        <Column field="fileSize" header="修改时间" style="width: 180px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">
              {{ dayjs(slotProps.data.mtime).format('YYYY-MM-DD HH:mm:ss') }}
            </div>
          </template>
        </Column>
        <Column
          field="count"
          header="出现次数"
          style="width: 120px; font-weight: bold"
          sortable
        ></Column>
        <Column field="examples" header="示例">
          <template #body="slotProps">
            <div>
              <div
                class="m-b-12px"
                contenteditable
                v-for="item in slotProps.data.examples.slice(0, 2)"
                v-html="item"
              ></div>
              <div
                v-if="slotProps.data.examples.length > 2"
                class="color-blue cursor-pointer text-center text-12px"
                @click="showMoreData(slotProps.data)"
              >
                查看全部
              </div>
            </div>
          </template></Column
        >
      </DataTable>
      <div
        class="flex items-center justify-center"
        v-if="!pptTableData.length"
        style="height: calc(100vh - 200px)"
      >
        <img src="./assets/images/empty1.png" alt="" />
      </div>
    </div>
    <div class="m-t-12px" v-if="showFileType === 'docx'">
      <DataTable
        v-if="wordTableData.length"
        :value="wordTableData"
        tableStyle="min-width: 50%"
        showGridlines
        selectionMode="single"
        scrollable
        stripedRows
        scrollHeight="calc(100vh - 200px)"
      >
        <!-- <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.data.id }}
          </template></Column
        > -->
        <Column field="fileName" header="文件名" style="font-weight: 500; width: 100px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              contenteditable
              v-html="slotProps.data.fileName"
            ></div> </template
        ></Column>
        <Column field="fileName" header="文件路径" style="width: 200px">
          <template #body="slotProps">
            <div
              class="m-b-12px color-blue cursor-pointer"
              @click="openLink(slotProps.data.filePath)"
              v-html="slotProps.data.filePath"
            ></div>
          </template>
        </Column>
        <Column field="fileSize" header="文件大小" style="width: 100px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">{{ formatFileSize(slotProps.data.fileSize) }}</div>
          </template>
        </Column>
        <Column field="fileSize" header="修改时间" style="width: 180px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">
              {{ dayjs(slotProps.data.mtime).format('YYYY-MM-DD HH:mm:ss') }}
            </div>
          </template>
        </Column>
        <Column
          field="count"
          header="出现次数"
          style="width: 120px; font-weight: bold"
          sortable
        ></Column>
        <Column field="examples" header="示例">
          <template #body="slotProps">
            <div>
              <div
                class="m-b-12px"
                contenteditable
                v-for="item in slotProps.data.examples.slice(0, 2)"
                v-html="item"
              ></div>
              <div
                v-if="slotProps.data.examples.length > 2"
                class="color-blue cursor-pointer text-center text-12px"
                @click="showMoreData(slotProps.data)"
              >
                查看全部
              </div>
            </div>
          </template></Column
        >
      </DataTable>
      <div
        class="flex items-center justify-center"
        v-if="!wordTableData.length"
        style="height: calc(100vh - 200px)"
      >
        <img src="./assets/images/empty1.png" alt="" />
      </div>
    </div>
    <div class="m-t-12px" v-if="showFileType === 'pdf'">
      <DataTable
        v-if="pdfTableData.length"
        :value="pdfTableData"
        tableStyle="min-width: 50%"
        showGridlines
        selectionMode="single"
        scrollable
        stripedRows
        scrollHeight="calc(100vh - 200px)"
      >
        <!-- <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.index + 1 }}
          </template></Column
        > -->
        <Column field="fileName" header="文件名" style="font-weight: 500; width: 100px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              contenteditable
              v-html="slotProps.data.fileName"
            ></div> </template
        ></Column>
        <Column field="fileName" header="文件路径" style="width: 200px">
          <template #body="slotProps">
            <div
              class="m-b-12px color-blue cursor-pointer"
              @click="openLink(slotProps.data.filePath)"
              v-html="slotProps.data.filePath"
            ></div>
          </template>
        </Column>
        <Column field="fileSize" header="文件大小" style="width: 100px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">{{ formatFileSize(slotProps.data.fileSize) }}</div>
          </template>
        </Column>
        <Column field="fileSize" header="修改时间" style="width: 180px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">
              {{ dayjs(slotProps.data.mtime).format('YYYY-MM-DD HH:mm:ss') }}
            </div>
          </template>
        </Column>
        <Column
          field="count"
          header="出现次数"
          style="width: 120px; font-weight: bold"
          sortable
        ></Column>
        <Column field="examples" header="示例">
          <template #body="slotProps">
            <div>
              <div
                class="m-b-12px"
                contenteditable
                v-for="item in slotProps.data.examples.slice(0, 2)"
                v-html="item"
              ></div>
              <div
                v-if="slotProps.data.examples.length > 2"
                class="color-blue cursor-pointer text-center text-12px"
                @click="showMoreData(slotProps.data)"
              >
                查看全部
              </div>
            </div>
          </template></Column
        >
      </DataTable>
      <div
        class="flex items-center justify-center"
        v-if="!pdfTableData.length"
        style="height: calc(100vh - 200px)"
      >
        <img src="./assets/images/empty1.png" alt="" />
      </div>
    </div>
    <div class="m-t-12px" v-if="showFileType === 'xlsx'">
      <DataTable
        v-if="excelTableData.length"
        :value="excelTableData"
        tableStyle="min-width: 50%"
        scrollable
        selectionMode="single"
        rowGroupMode="rowspan"
        groupRowsBy="filePath"
        scrollHeight="calc(100vh - 200px)"
      >
        <!-- <Column field="index" header="序号" style="width: 70px">
          <template #body="slotProps">
            {{ slotProps.index + 1 }}
          </template></Column
        > -->
        <Column field="fileName" header="文件名" style="font-weight: 500; width: 100px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              contenteditable
              v-html="slotProps.data.fileName"
            ></div> </template
        ></Column>
        <Column field="filePath" header="文件路径" style="width: 200px">
          <template #body="slotProps">
            <div
              class="m-b-12px cursor-pointer"
              @click="openLink(slotProps.data.filePath)"
              v-html="slotProps.data.filePath"
            ></div>
          </template>
        </Column>

        <Column field="sheetName" header="子表名"></Column>
        <Column field="fileSize" header="文件大小" style="width: 100px">
          <template #body="slotProps">
            <div class="color-#676767 text-14px">{{ formatFileSize(slotProps.data.fileSize) }}</div>
          </template>
        </Column>
        <Column field="fileSize" header="修改时间" style="width: 180px">
          <template #body="slotProps">
            <div class="color-#888 text-14px">
              {{ dayjs(slotProps.data.mtime).format('YYYY-MM-DD HH:mm:ss') }}
            </div>
          </template>
        </Column>
        <Column field="count" header="出现次数" style="width: 120px; font-weight: bold" sortable>
        </Column>
        <Column field="examples" header="示例">
          <template #body="slotProps">
            <div>
              <div
                class="m-b-12px"
                contenteditable
                v-for="item in slotProps.data.examples.slice(0, 2)"
                v-html="item"
              ></div>
              <div
                v-if="slotProps.data.examples.length > 2"
                class="color-blue cursor-pointer text-center text-12px"
                @click="showMoreData(slotProps.data)"
              >
                查看全部
              </div>
            </div>
          </template></Column
        >
      </DataTable>
      <div
        class="flex items-center justify-center"
        v-if="!excelTableData.length"
        style="height: calc(100vh - 200px)"
      >
        <img src="./assets/images/empty1.png" alt="" />
      </div>
    </div>
  </div>
</template>
