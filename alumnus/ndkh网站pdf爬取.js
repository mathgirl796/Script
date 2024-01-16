

async function parseImgSrc(targetUrl) {
    var imgSrcList = []
    var title
    // 发起GET请求
    await fetch(targetUrl).then(response=>response.text()).then(html=>{
        // 创建一个虚拟的HTML元素，将返回的HTML插入其中
        const virtualDOM = document.createElement('html');
        virtualDOM.innerHTML = html;
        // 获取所有图片元素
        imgElements = virtualDOM.querySelectorAll('img');
        title = virtualDOM.querySelector('h3').innerText.replace(/\s*/g,"")
        folder_title = document.querySelector("body > div.component > div > table.sticky-thead.table.table-striped.table-bordered > thead > tr:nth-child(1) > th:nth-child(2)").innerText
        console.log(folder_title)
    }).catch(error=>console.error('Error fetching data:', error));
    // 遍历图片元素并将图片链接存入imgSrcList
    for (let index = 0; index < imgElements.length; index++) {
        imgSrcList[index] = imgElements[index].getAttribute('data-src');
    }
    return {imgSrcList:imgSrcList, title:title}
}

async function loadJspdf() {
    // 加载jspdf库
    var script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    document.head.appendChild(script);
    var promise = new Promise((reslove)=>{
        script.onload = async function() {
            console.log("jdpdf loaded")
            reslove();
        }
    }
    )
    await promise
}


async function runJspdf(imgSrcList, title) {
    //创建pdf文件
    const {jsPDF} = window.jspdf;
    const doc = new jsPDF();
    doc.deletePage(1);
    
    //遍历图片加入pdf文件
    for (index = 0; index < imgSrcList.length; index++) {
        console.log("addImage", title, index)
        var img = new Image()
        img.src = imgSrcList[index]
        var width, height
        var promise = new Promise((reslove)=>{
            img.onload = async function() {
                width = img.width;
                height = img.height;
                reslove();
            }
        }
        )
        await promise
        doc.addPage([width, height], "p")
        doc.addImage(img, 'JPEG', 0, 0, width, height)
    }
    
    //保存pdf文件
    doc.save(title + ".pdf");
}

async function downloadReport(reportWorkId, type, underline=0) {
    var url = `https://ndkh.hit.edu.cn/reportWork_show?id=${reportWorkId}&isPlan=${type}&_=${underline}`
    var obj = await parseImgSrc(url)
    await runJspdf(obj.imgSrcList, obj.title)
}

async function main() {
    await loadJspdf()
    table = document.querySelector(".table")
    works = table.querySelectorAll("a")
    console.log("detect", works.length, "works")
    for (let index=0;index<works.length;index++){
        var reportWorkId = works[index].getAttribute('onclick').split("'")[1]
        var type = works[index].getAttribute('onclick').split(",").length > 1 ? 1:0;
        await downloadReport(reportWorkId, type)
        console.log(`progress: ${index+1}/${works.length}`)
    }
    return "main done"
}

await main()