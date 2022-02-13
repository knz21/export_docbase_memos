const docBaseAccessToken: string = PropertiesService.getScriptProperties().getProperty('docbase_access_token')
const sheetId = ''
const imagesFolderId = ''
const docbaseTeamId = ''
const docbaseUserId = ''

const doExport = () => {
    const memos = fetchDocbaseMemo(`https://api.docbase.io/teams/dely/posts?q=author:${docbaseUserId}`)
    if (memos.length <= 0) return
    Logger.log(`Memos: ${memos.length}`)

    const sheet = SpreadsheetApp.openById(sheetId)
    const memoSheet = sheet.getSheetByName('memo') || sheet.insertSheet('memo')
    memoSheet.getRange(1, 1, memos.length, 3).setValues(memos.map(memo => [memo.id, memo.title, memo.body]))
    memoSheet.setRowHeightsForced(1, memos.length, 21);
}

const fetchDocbaseMemo = (url: string, memos: Memo[] = []): Memo[] => {
    const docbaseApiHeaders = {
        'Content-type': 'application/json',
        'X-DocBaseToken': docBaseAccessToken
    }
    const params: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        headers: docbaseApiHeaders,
        method: 'get'
    }

    const json = JSON.parse(UrlFetchApp.fetch(url, params).getContentText())
    const newMemos = memos.concat(json.posts.map(post => {
        return {
            id: post.id,
            title: post.title,
            body: post.body
        }
    }))
    const newUrl = json.meta.next_page

    if (newUrl) {
        return fetchDocbaseMemo(newUrl, newMemos)
    } else {
        return newMemos
    }
}

const downloadMemoImages = () => {
    const sheet = SpreadsheetApp.openById(sheetId)
    const sheetNames = sheet.getSheets().map(sheet => sheet.getName())
    const memoSheet = sheet.getSheetByName('memo')
    const memos = memoSheet.getRange(1, 1, memoSheet.getLastRow(), 3).getValues()
    const targetMemos = memos.filter(memo => !sheetNames.includes(memo[0].toString()))
    const parentImageFolder = DriveApp.getFolderById(imagesFolderId)
    Logger.log(`Target memos: ${targetMemos.length}`)
    targetMemos.forEach(memo => {
        const memoId = memo[0]
        const title = memo[1]
        const body = memo[2]
        const urls = body.match(/https?:\/\/[^ \)\]|\\r\\n]*\.(?:png|jpg)/g)
        if (urls) {
            Logger.log(`${memoId}: ${urls.length} images (${title})`)
            const existingImageFolder = parentImageFolder.getFoldersByName(memoId)
            while (existingImageFolder.hasNext()) {
                Logger.log('Delete existing folder')
                existingImageFolder.next().setTrashed(true)
            }
            const newImageFolder = parentImageFolder.createFolder(memoId)
            const urlsList = []

            urls.forEach(url => {
                const blob = getDocbaseImageBlob(url)
                const imageFileInRoot = DriveApp.createFile(blob.setName(url))
                const imageFile = imageFileInRoot.makeCopy(imageFileInRoot.getName(), newImageFolder)
                imageFileInRoot.setTrashed(true)
                urlsList.push([url, imageFile.getUrl()])
                Logger.log(`Downloaded: ${url}`)
            })

            const imageSheet = sheet.insertSheet().setName(memoId)
            imageSheet.getRange(1, 1, urlsList.length, 2).setValues(urlsList)
        } else {
            Logger.log(`${memoId}: No images (${title})`)
        }
    })
}

const getDocbaseImageBlob = (url: string) => {
    const headers = {
        cookie: `team_ids=${docbaseTeamId};`,
        'X-DocBaseToken': docBaseAccessToken
    }
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        headers: headers,
        method: 'get',
        muteHttpExceptions: true
    }
    return UrlFetchApp.fetch(url, options).getBlob()
}

const deleteImageSheets = () => {
    const sheet = SpreadsheetApp.openById(sheetId)
    sheet.getSheets().forEach(childSheet => {
        if (childSheet.getName() != 'memo') {
            sheet.deleteSheet(childSheet)
        }
    })
}

interface Memo {
    id: string,
    title: string,
    body: string
}