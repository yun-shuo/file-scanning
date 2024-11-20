const Store = require('electron-store')
const store = new Store()
let schemeList = null
export const addScheme = (event, dataString) => {
  const obj = JSON.parse(dataString)
  const schemeList = getSchemeList()
  schemeList.unshift(obj)
  store.set('schemeList', JSON.stringify(schemeList))
  return true
}
export const getSchemeList = () => {
  if (schemeList) {
    return schemeList
  } else {
    schemeList = store.get('schemeList') ? JSON.parse(store.get('schemeList')) : []
  }
  return schemeList
}
