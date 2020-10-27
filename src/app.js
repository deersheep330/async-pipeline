const axios = require('axios')
const qs = require('querystring')
const FormData = require('form-data')
const fs = require('fs')
const CircularJSON = require('circular-json')

const { to } = require('./utils/to.js')
const { sleep } = require('./utils/sleep.js')

let threadNumber = 0
let timeLimit = 1000 * 60 * 30
const baseUrl = 'https://app-jp.patentcloud.com'

var args = process.argv.slice(2)
const account = args[0] || ''
const password = args[1] || ''
const filename = args[2] || ''

let workspaceId = null

let startTime = new Date()
let uploadTime = null
let cs1Time = null
let vh3Time = null

let sendLog = (log) => {
    console.log(log)
}

let sendError = (err) => {
    console.log(err)
}

sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] start for ' + filename)

const defaultHeaders = {
    'Referer': baseUrl + '/',
    'Accept-Language': 'zh-TW,zh;q=0.8,en-US;q=0.5,en;q=0.3',
    'Accept-Encoding': 'gzip, deflate, br',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:66.0) Gecko/20100101 Firefox/66.0',
    'Accept': 'application/json, text/plain, */*'
}

const regexForCookie = /(?<key>[^=]+)=(?<value>[^\;]*);?\s?/g

// step 1
const login = async () => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['Content-Type']	= 'application/x-www-form-urlencoded'

        let body = {
            'account': account,
            'password': password,
            'rememberMe': 'true',
            'forceLogin[isTrusted]': 'true',
            'captcha': ''
        }

        let url = baseUrl + '/member/login.do'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            data: qs.stringify(body),
            headers: headers
        }))

        if (err) { reject(err) }
        else {
            let parsedCookie = {}
            for (let str of res.headers['set-cookie']) {
                let matched = str.matchAll(regexForCookie)
                for (let m of matched) {
                    let { key, value } = m.groups
                    parsedCookie[key] = value
                }
            }
            let cookieToBeUsed = ''
            cookieToBeUsed += 't=' + parsedCookie['t'] + '; '
            cookieToBeUsed += 'v=' + parsedCookie['v'] + '; '
            cookieToBeUsed += 'SESSION=' + parsedCookie['SESSION'] + '; '
            resolve(cookieToBeUsed) 
        }

    })
}

// step 2
const getUserInfo = async (cookie) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie

        let url = baseUrl + '/member/login/userInfo.do'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))
        
        if (err) { reject(err) }
        else { resolve(res.data.user) }

    })
}

// step 3
const uploadFile = async (cookie) => {
    return new Promise( async (resolve, reject) => {

        // step 1: upload file

        const form = new FormData()
        form.append('file', fs.createReadStream(__dirname + '/files/' + filename))

        const inqHeaders = {
            'X-Inq-PcMemberId':	'1fe68bba-4cba-4b8d-8f9b-064dff998ae6',
            'X-Inq-Account': account,
            'X-Inq-Credential':	'9c2e2b70-4bdf-4eec-a2ea-614397381de4',
            'X-Inq-Session': '869b7ccb4792010008e0eca06425df59'
        }

        let headers = form.getHeaders()
        headers = {
            ...headers,
            ...defaultHeaders,
            ...inqHeaders,
            cookie: cookie
        }

        let url = baseUrl + '/v2/file/upload'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            data: form,
            headers: headers
        }))

        let workspaceId = ''
        if (err) { reject(err) }
        else if (res.data.result !== 'SUCCESS') { reject(res.data) }
        else { workspaceId = res.data.data.workspaceId }

        // step 2: check status
        let lastRes = {}
        let _status = 2
        let _headers = JSON.parse(JSON.stringify(defaultHeaders))
        let _params = { workspaceId: workspaceId }
        let _url = baseUrl + '/v2/completenesshealthcheck'
        while (_status === 2) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'get',
                url: _url,
                params: _params,
                headers: _headers
            }))

            // ignore any error, keep polling

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('COMPLETENESS_PROGRESS_STATUS')) {
                _status = res.data.data.COMPLETENESS_PROGRESS_STATUS
                sendLog('[' +  (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] upload file status = ' + _status)
                lastRes = res
            }
        }

        if (_status !== 0) reject(CircularJSON.stringify(lastRes.data))

        resolve(workspaceId)

    })
}

// step 4
const modelPrivilege = async (cookie) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie

        let url = baseUrl + '/member/privilege/modelPrivilege.do'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else { resolve(res.data) }

    })
}

// step 5
const verifyWorkspace = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId

        let url = baseUrl + '/v2/order/verify'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else if (res.data.result !== 'SUCCESS') { reject(CircularJSON.stringify(res.data)) }
        else { resolve(res.data.message) }

    })
}

// step 6
const getWorkspaceSetting = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId

        let url = baseUrl + '/v2/chart/workspaceSetting'

        const [err, res] = await to(axios({
            method: 'get',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else if (res.data.result !== 'SUCCESS') { reject(CircularJSON.stringify(res.data)) }
        else { resolve(res.data.data) }

    })
}

// step 7
const A1 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie
        headers['Content-Type']	= 'application/json; charset=utf-8'
        headers['x-inq-workspaceid'] = workspaceId

        let url = baseUrl + '/v2/chart/A_1'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else { resolve(res.data) }

    })
}

// step 8
const A2 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie
        headers['Content-Type']	= 'application/json; charset=utf-8'
        headers['x-inq-workspaceid'] = workspaceId

        let url = baseUrl + '/v2/chart/A_2'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else { resolve(res.data) }

    })
}

// step 9
const setTrial = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId

        let url = baseUrl + '/v2/account/trial'

        const [err, res] = await to(axios({
            method: 'post',
            url: url,
            headers: headers
        }))

        if (err) { reject(err) }
        else if (res.data.result !== 'SUCCESS') { reject(CircularJSON.stringify(res.data)) }
        else { resolve(res.data.message) }

    })
}

const cs1 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let lastRes
        let chartData = {}
        let progress = 9
        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['Content-Type'] = 'application/json; charset=UTF-8'
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId
        let url = baseUrl + '/v2/chart/CS_1'
        while (progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                url: url,
                headers: headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('chart') && res.data.data.chart.hasOwnProperty('progress')) {
                progress = res.data.data.chart.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] cs1 progress = ' + progress)
                if (progress === 1) chartData = res.data.data.chart
            }

            lastRes = res
        }

        if (progress !== 1) {
            if (lastRes && lastRes.data) reject(CircularJSON.stringify(lastRes.data))
            else reject(progress)
        }

        resolve(chartData)

    })
}

const vh1 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let lastRes = {}

        let progress = 9
        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['Content-Type'] = 'application/json; charset=UTF-8'
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId
        let data = { countBy: 2 }
        let url = baseUrl + '/v2/chart/topNList/VH_1'
        while (progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: data,
                url: url,
                headers: headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('progress')) {
                progress = res.data.data.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh1 top n list progress = ' + progress)
            }

            lastRes = res
        }

        if (progress !== 1) reject(CircularJSON.stringify(lastRes.data))

        let chartData = {}
        let _progress = 9
        let _headers = JSON.parse(JSON.stringify(defaultHeaders))
        _headers['Content-Type'] = 'application/json; charset=UTF-8'
        _headers['cookie'] = cookie
        _headers['x-inq-workspaceid'] = workspaceId

        let _data = ''
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] calc vh1 for ' + filename)
        switch (filename) {
            case '200.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"EP\",\"WO\",\"CN\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"TW\",\"GB\",\"EP\",\"ES\",\"AU\",\"WO\",\"US\",\"IT\",\"DE\",\"KR\",\"RU\",\"CN\",\"JP\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"PANASONIC CORP\",\"OSTERHOUT GROUP INC\",\"PANASONIC INTELLECTUAL PROPERTY MANAGEMENT CO LTD\",\"SONY CORP\",\"NEXEON LTD\",\"QUALCOMM INC\",\"PANASONIC ELECTRIC WORKS CO LTD\",\"CANON KK\",\"MENTOR ACQUISITION ONE LLC\",\"TOSHIBA KK\"]}}}","countBy":2}
                break;
            case '3000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"JP\",\"EP\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"GR\",\"NO\",\"SG\",\"IT\",\"NL\",\"RU\",\"US\",\"GB\",\"ES\",\"WO\",\"AU\",\"AT\",\"EP\",\"FR\",\"CN\",\"JP\",\"DE\",\"KR\",\"TW\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"IBM CORP\",\"MICROSOFT CORP\",\"CISCO TECHNOLOGY INC\",\"AT&T INTELLECTUAL PROPERTY I LP\",\"MICROSOFT TECHNOLOGY LICENSING LLC\",\"APPLE INC\",\"AMAZON TECHNOLOGIES INC\",\"QUALCOMM INC\",\"SAMSUNG ELECTRONICS CO LTD\",\"GOOGLE LLC\"]}}}","countBy":2}
                break;
            case '45000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"NZ\",\"SE\",\"IT\",\"IL\",\"PT\",\"PL\",\"AU\",\"GB\",\"KR\",\"EP\",\"JP\",\"CN\",\"US\",\"ES\",\"SG\",\"CA\",\"RU\",\"BR\",\"MX\",\"NL\",\"HK\",\"MY\",\"EM\",\"AT\",\"DE\",\"HU\",\"DK\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"PT\",\"GR\",\"GB\",\"AU\",\"CN\",\"JP\",\"KR\",\"BE\",\"RU\",\"IT\",\"US\",\"ES\",\"EP\",\"TW\",\"CH\",\"WO\",\"FR\",\"NL\",\"SG\",\"AT\",\"NO\",\"DE\",\"CZ\",\"DK\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"SHARP CORP\",\"SAMSUNG ELECTRONICS CO LTD\",\"CANON KK\",\"SEMICONDUCTOR ENERGY LABORATORY CO LTD\",\"RICOH CO LTD\",\"SAMSUNG DISPLAY CO LTD\",\"SONY CORP\",\"LG ELECTRONICS INC\",\"TOSHIBA KK\",\"BOE TECHNOLOGY GROUP CO LTD\"]}}}","countBy":2}
                break;
            default:
                break;
        }
        
        let _url = baseUrl + '/v2/chart/VH_1'
        while (_progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: _data,
                url: _url,
                headers: _headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('chart') && res.data.data.chart.hasOwnProperty('progress')) {
                _progress = res.data.data.chart.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh1 chart progress = ' + _progress)
                if (_progress === 1) chartData = res.data.data.chart
            }

            lastRes = res
        }

        if (_progress !== 1) {
            if (lastRes && lastRes.data) reject(CircularJSON.stringify(lastRes.data))
            else reject(_progress)
        }

        resolve(chartData)
    })
}

const vh2 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let lastRes = {}

        let progress = 9
        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['Content-Type'] = 'application/json; charset=UTF-8'
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId
        let data = { countBy: 2 }
        let url = baseUrl + '/v2/chart/topNList/VH_2'
        while (progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: data,
                url: url,
                headers: headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('progress')) {
                progress = res.data.data.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh2 top n list progress = ' + progress)
            }

            lastRes = res
        }

        if (progress !== 1) reject(CircularJSON.stringify(lastRes.data))

        let chartData = {}
        let _progress = 9
        let _headers = JSON.parse(JSON.stringify(defaultHeaders))
        _headers['Content-Type'] = 'application/json; charset=UTF-8'
        _headers['cookie'] = cookie
        _headers['x-inq-workspaceid'] = workspaceId

        let _data = ''
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] calc vh2 for ' + filename)
        switch (filename) {
            case '200.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"EP\",\"WO\",\"CN\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"TW\",\"GB\",\"EP\",\"ES\",\"AU\",\"WO\",\"US\",\"IT\",\"DE\",\"KR\",\"RU\",\"CN\",\"JP\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"PANASONIC CORP\",\"OSTERHOUT GROUP INC\",\"PANASONIC INTELLECTUAL PROPERTY MANAGEMENT CO LTD\",\"SONY CORP\",\"NEXEON LTD\",\"QUALCOMM INC\",\"PANASONIC ELECTRIC WORKS CO LTD\",\"CANON KK\",\"MENTOR ACQUISITION ONE LLC\",\"TOSHIBA KK\"]}}}","countBy":2}
                break;
            case '3000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"JP\",\"EP\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"GR\",\"NO\",\"SG\",\"IT\",\"NL\",\"RU\",\"US\",\"GB\",\"ES\",\"WO\",\"AU\",\"AT\",\"EP\",\"FR\",\"CN\",\"JP\",\"DE\",\"KR\",\"TW\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"IBM CORP\",\"MICROSOFT CORP\",\"CISCO TECHNOLOGY INC\",\"AT&T INTELLECTUAL PROPERTY I LP\",\"MICROSOFT TECHNOLOGY LICENSING LLC\",\"APPLE INC\",\"AMAZON TECHNOLOGIES INC\",\"QUALCOMM INC\",\"SAMSUNG ELECTRONICS CO LTD\",\"GOOGLE LLC\"]}}}","countBy":2}
                break;
            case '45000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"NZ\",\"SE\",\"IT\",\"IL\",\"PT\",\"PL\",\"AU\",\"GB\",\"KR\",\"EP\",\"JP\",\"CN\",\"US\",\"ES\",\"SG\",\"CA\",\"RU\",\"BR\",\"MX\",\"NL\",\"HK\",\"MY\",\"EM\",\"AT\",\"DE\",\"HU\",\"DK\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"PT\",\"GR\",\"GB\",\"AU\",\"CN\",\"JP\",\"KR\",\"BE\",\"RU\",\"IT\",\"US\",\"ES\",\"EP\",\"TW\",\"CH\",\"WO\",\"FR\",\"NL\",\"SG\",\"AT\",\"NO\",\"DE\",\"CZ\",\"DK\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"SHARP CORP\",\"SAMSUNG ELECTRONICS CO LTD\",\"CANON KK\",\"SEMICONDUCTOR ENERGY LABORATORY CO LTD\",\"RICOH CO LTD\",\"SAMSUNG DISPLAY CO LTD\",\"SONY CORP\",\"LG ELECTRONICS INC\",\"TOSHIBA KK\",\"BOE TECHNOLOGY GROUP CO LTD\"]}}}","countBy":2}
                break;
            default:
                break;
        }

        let _url = baseUrl + '/v2/chart/VH_2'
        while (_progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: _data,
                url: _url,
                headers: _headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('chart') && res.data.data.chart.hasOwnProperty('progress')) {
                _progress = res.data.data.chart.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh2 chart progress = ' + _progress)
                if (_progress === 1) chartData = res.data.data.chart
            }

            lastRes = res
        }

        if (_progress !== 1) {
            if (lastRes && lastRes.data) reject(CircularJSON.stringify(lastRes.data))
            else reject(_progress)
        }

        resolve(chartData)
    })
}

const vh3 = async (cookie, workspaceId) => {
    return new Promise( async (resolve, reject) => {

        let lastRes = {}

        let progress = 9
        let headers = JSON.parse(JSON.stringify(defaultHeaders))
        headers['Content-Type'] = 'application/json; charset=UTF-8'
        headers['cookie'] = cookie
        headers['x-inq-workspaceid'] = workspaceId
        let data = { countBy: 2 }
        let url = baseUrl + '/v2/chart/topNList/VH_3'
        while (progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: data,
                url: url,
                headers: headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('progress')) {
                progress = res.data.data.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh3 top n list progress = ' + progress)
            }

            lastRes = res
        }

        if (progress !== 1) reject(CircularJSON.stringify(lastRes.data))

        let chartData = {}
        let _progress = 9
        let _headers = JSON.parse(JSON.stringify(defaultHeaders))
        _headers['Content-Type'] = 'application/json; charset=UTF-8'
        _headers['cookie'] = cookie
        _headers['x-inq-workspaceid'] = workspaceId

        let _data = ''
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] calc vh3 for ' + filename)
        switch (filename) {
            case '200.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"EP\",\"WO\",\"CN\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"TW\",\"GB\",\"EP\",\"ES\",\"AU\",\"WO\",\"US\",\"IT\",\"DE\",\"KR\",\"RU\",\"CN\",\"JP\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"PANASONIC CORP\",\"OSTERHOUT GROUP INC\",\"PANASONIC INTELLECTUAL PROPERTY MANAGEMENT CO LTD\",\"SONY CORP\",\"NEXEON LTD\",\"QUALCOMM INC\",\"PANASONIC ELECTRIC WORKS CO LTD\",\"CANON KK\",\"MENTOR ACQUISITION ONE LLC\",\"TOSHIBA KK\"]}}}","countBy":2,"stackByRank":1}
                break;
            case '3000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"JP\",\"EP\",\"US\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"GR\",\"NO\",\"SG\",\"IT\",\"NL\",\"RU\",\"US\",\"GB\",\"ES\",\"WO\",\"AU\",\"AT\",\"EP\",\"FR\",\"CN\",\"JP\",\"DE\",\"KR\",\"TW\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"IBM CORP\",\"MICROSOFT CORP\",\"CISCO TECHNOLOGY INC\",\"AT&T INTELLECTUAL PROPERTY I LP\",\"MICROSOFT TECHNOLOGY LICENSING LLC\",\"APPLE INC\",\"AMAZON TECHNOLOGIES INC\",\"QUALCOMM INC\",\"SAMSUNG ELECTRONICS CO LTD\",\"GOOGLE LLC\"]}}}","countBy":2,"stackByRank":1}
                break;
            case '45000.xlsx':
                _data = {"queryString":"{\"portfolio\":{\"country\":{\"$in\":[\"NZ\",\"SE\",\"IT\",\"IL\",\"PT\",\"PL\",\"AU\",\"GB\",\"KR\",\"EP\",\"JP\",\"CN\",\"US\",\"ES\",\"SG\",\"CA\",\"RU\",\"BR\",\"MX\",\"NL\",\"HK\",\"MY\",\"EM\",\"AT\",\"DE\",\"HU\",\"DK\",\"FR\"]},\"statusCode\":{\"$in\":[1,2,3]}},\"citation\":{\"country\":{\"$in\":[\"PT\",\"GR\",\"GB\",\"AU\",\"CN\",\"JP\",\"KR\",\"BE\",\"RU\",\"IT\",\"US\",\"ES\",\"EP\",\"TW\",\"CH\",\"WO\",\"FR\",\"NL\",\"SG\",\"AT\",\"NO\",\"DE\",\"CZ\",\"DK\"]},\"statusCode\":{\"$in\":[1,2,3,4,5,6,8]},\"statusComp\":8,\"docdbAssigneesNormalFacetname\":{\"$in\":[\"SHARP CORP\",\"SAMSUNG ELECTRONICS CO LTD\",\"CANON KK\",\"SEMICONDUCTOR ENERGY LABORATORY CO LTD\",\"RICOH CO LTD\",\"SAMSUNG DISPLAY CO LTD\",\"SONY CORP\",\"LG ELECTRONICS INC\",\"TOSHIBA KK\",\"BOE TECHNOLOGY GROUP CO LTD\"]}}}","countBy":2,"stackByRank":1}
                break;
            default:
                break;
        }
        
        let _url = baseUrl + '/v2/chart/VH_3'
        while (_progress === 9) {

            await sleep(2000)

            const [err, res] = await to(axios({
                method: 'post',
                data: _data,
                url: _url,
                headers: _headers
            }))

            if (res && res.hasOwnProperty('data') && res.data.hasOwnProperty('data') && res.data.data.hasOwnProperty('chart') && res.data.data.chart.hasOwnProperty('progress')) {
                _progress = res.data.data.chart.progress
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh3 chart progress = ' + _progress)
                if (_progress === 1) chartData = res.data.data.chart
            }

            laastRes = res
        }
/*
        if (_progress !== 1) {
            if (lastRes && lastRes.data) reject(CircularJSON.stringify(lastRes.data))
            else reject(_progress)
        }
*/
        resolve(chartData)
    })
}

(async () => {

    try {
        let cookie = await login()
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] login and get cookie: ' + cookie)
        await sleep(2000)

        const userInfo = await getUserInfo(cookie)
        await sleep(2000)

        // upload file start
        const uploadFileStartTime = Date.now()
        workspaceId = await uploadFile(cookie)
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] upload file and get workspace id = ' + workspaceId)
        await sleep(5000)

        uploadTime = Date.now() - uploadFileStartTime
        // upload file end

        let modelPrivilegeResult
        try {
            modelPrivilegeResult = await modelPrivilege(cookie)
        }
        catch (e) {
            sendError(JSON.stringify(e) + ', wordspace id = ' + workspaceId)
            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] model privilege failed ' + JSON.stringify(e))

            cookie = await login()
            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] relogin and get cookie: ' + cookie)
            await sleep(2000)

            modelPrivilegeResult = await modelPrivilege(cookie)
        }
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] model privilege ok')
        await sleep(5000)

        let verifyWorkspaceResult
        try {
            verifyWorkspaceResult = await verifyWorkspace(cookie, workspaceId)
        }
        catch (e) {
            sendError(JSON.stringify(e) + ', wordspace id = ' + workspaceId)
            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] verify workspace failed ' + JSON.stringify(e))

            cookie = await login()
            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] relogin and get cookie: ' + cookie)
            await sleep(2000)

            verifyWorkspaceResult = await verifyWorkspace(cookie, workspaceId)
        }
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] verify workspace ok')
        await sleep(5000)

        let workspaceSetting = await getWorkspaceSetting(cookie, workspaceId)
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] get workspace setting ok')
        await sleep(5000)

        let A1Res = await A1(cookie, workspaceId)
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] request for A1 ok')
        await sleep(5000)

        let A2Res = await A2(cookie, workspaceId)
        sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] request for A2 ok')
        await sleep(5000)

        let retry = 0, maxRetry = 5
        while (retry++ < maxRetry) {
            try {
                let setTrialResult = await setTrial(cookie, workspaceId)
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] set trial success')
                break
            }
            catch (e) {
                console.log(JSON.stringify(e) + ', wordspace id = ' + workspaceId)
                sendError(JSON.stringify(e) + ', wordspace id = ' + workspaceId)
            }
            await sleep(5000)
        }

        // get chart start
        const getChartStartTime = Date.now()

        let cs1Done = false, vh3Done = false

            cs1(cookie, workspaceId).then((chartData) => {
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] cs1 complete')
                cs1Time = Date.now() - getChartStartTime
            }).catch ((err) => {
                sendError(err)
            }).finally(() => {
                cs1Done = true
            })

            vh3(cookie, workspaceId).then((chartData) => {
                sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] vh3 complete')
                vh3Time = Date.now() - getChartStartTime
            }).catch ((err) => {
                sendError(err)
            }).finally(() => {
                vh3Done = true
            })

            while(!cs1Done || !vh3Done) { await sleep(2000) }

            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] complete one task loop')

        } catch (err) {
            sendError(err)
            workspaceId = null
        } finally {
            sendLog('[' + (new Date()).toLocaleTimeString() + '][thread ' + threadNumber + '] ' + uploadTime / 1000 + ',' + cs1Time / 1000 + ',' + vh3Time / 1000)
            fs.appendFileSync(filename.replace('xlsx', 'csv'), startTime + ',' + uploadTime / 1000 + ',' + cs1Time / 1000 + ',' + vh3Time / 1000 + '\n')
        }

})()
