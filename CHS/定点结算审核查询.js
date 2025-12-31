idlist=["220303196307032458",
"220322195205105315"]

myid = "220303196804023640"
function myreset(){return new Promise(resolve=>{
//重置表单
setTimeout(()=>{
document.querySelector("#el-collapse-content-1749 > div > form > div > div.el-row__footbar > button:nth-child(2)").click()
console.log(0,"已点击重置")
resolve()
},0)
})}

function myinput(){return new Promise(resolve=>{
//输入身份证号
setTimeout(()=>{
document.querySelector("#el-collapse-content-1749 > div > form > div > div:nth-child(2) > div > div > div > input").value=myid
document.querySelector("#el-collapse-content-1749 > div > form > div > div:nth-child(2) > div > div > div > input").dispatchEvent(new Event('input',{bubbles:true,cancelable:true}))
console.log(1,'['+myid+']')
resolve()
},500)
})}

function mypush(){return new Promise(resolve=>{
//点击查询按钮
setTimeout(()=>{
document.querySelector("#el-collapse-content-1749 > div > form > div > div.el-row__footbar > button.el-button.el-button--primary.el-button--medium").click()
console.log(2,"已点击查询")
resolve()
},500)
})}

function myget(){return new Promise(resolve=>{
//获取总页数信息
setTimeout(()=>{
a = document.querySelector("#el-collapse-content-747 > div > div > section > div.pagination-left.el-pagination.is-background > span").innerText
console.log(3,'['+a.trim()+']')
//获取当前页最后一条查询结果信息
a = document.querySelector("#el-collapse-content-747 > div > div > div > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody").lastElementChild.textContent
console.log(4,'['+a.trim()+']')
resolve()
},1000)
})}

function mypush2(){return new Promise(resolve=>{
//点击当前页最后一条的"查看明细"按钮
setTimeout(()=>{
document.querySelector("#el-collapse-content-747 > div > div > div > div.el-table__fixed-right > div.el-table__fixed-body-wrapper > table > tbody").lastElementChild.lastElementChild.querySelector('button').click()
console.log(5,"已点击查看明细")
resolve()
},500)
})}

function myget2(){return new Promise(resolve=>{
//获取明细信息
setTimeout(()=>{
a = document.querySelector("#app > div > div.el-dialog__wrapper > div > div.el-dialog__body.body-padding-top.body-padding-bottom > div > div > div.el-table__body-wrapper.is-scrolling-none > table > tbody").innerText
console.log(6,'['+a.trim()+']')
resolve()
},500)
})}

for (i=0;i<idlist.length;i++){
    console.log(i+'/'+idlist.length)
    myid=idlist[i]
    await myreset()
    .then(a=>myinput(a))
    .then(a=>mypush(a))
    .then(a=>myget(a))
    .then(a=>mypush2(a))
    .then(a=>myget2(a))
}

