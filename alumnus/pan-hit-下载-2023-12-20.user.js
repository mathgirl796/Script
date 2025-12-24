// ==UserScript==
// @name         pan-hit-下载
// @namespace    http://tampermonkey.net/
// @version      2025-12-24
// @description  try to take over the world!
// @author       You
// @match        panowa.hit.edu.cn/*
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @grant        none
// ==/UserScript==

/* 使用说明：把@match后面的服务器地址改成网盘里pdf页面中iframe元素的地址 */

function replaceNumberInString(inputString, replacementNumber) {
  /* 使用正则表达式匹配字符串中的 n=p1.img */
  const regex = /n=p(\d+)\.img/g;

  /* 使用 replace 方法进行替换 */
  const resultString = inputString.replace(regex, (match, group1) => {
    /* 将匹配到的数字部分替换为传入的整数 */
    return `n=p${replacementNumber}.img`;
  });

  return resultString;
}

async function loadJspdf() {
    /* 加载jspdf库 */
    var script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
    document.head.appendChild(script);
    var promise = new Promise((reslove)=>{
        script.onload = async function() {
            console.log("jspdf loaded");
            reslove();
        }
    }
    );
    await promise;
}

async function runJspdf() {
    var title = document.title; /* typeof title: string */
    console.log(title);
    var imgSrcTemplate = document.querySelector(".WACPageImg.WacAirSpaceShared_BasicClass").src;
    console.log(imgSrcTemplate);

    /* 创建pdf文件 */
    await loadJspdf();
    const {jsPDF} = window.jspdf;
    const doc = new jsPDF();
    doc.deletePage(1);

    /* 逐个尝试各个页码 */
    var i = 1;
    var stop = false;
    var width
    var height
    while (true) {
        var realImgSrc = replaceNumberInString(imgSrcTemplate, i);
        var img = new Image();
        img.src = realImgSrc;
        var promise = new Promise((reslove)=>{
            img.onload = async function() {
                width = img.width;
                height = img.height;
                reslove();
            }
            img.onerror = async function() {
                stop = true;
                reslove();
            }
        }
        )
        await promise;
        if (stop) {
            break;
        }
        else {
            console.log(width, height, promise) ;
            doc.addPage([width, height], "p");
            doc.addImage(img, 'JPEG', 0, 0, width, height);
        }
        i += 1;
    }

    /* 保存pdf文件 */
    doc.save(title);
}

function getIframeDocument(iframe) {
    return new Promise(resolve => {
        if (iframe.contentDocument) {
            resolve(iframe.contentDocument);
        } else {
            iframe.addEventListener('load', e => {
                resolve(iframe.contentDocument);
            });
        }
    });
}

(function() {
    'use strict';

    // Your code here...

    document.onreadystatechange = function () {
        if (document.readyState === "complete"&&window.top !== window.parent) {

            // 创建按钮元素
            var button = document.createElement("button");

            // 设置按钮文本
            button.innerHTML = "Click me!";

            // 设置按钮样式（可选）
            button.style.position = "fixed";
            button.style.top = "0";
            button.style.left = "0";
            button.style.zIndex = "9999";
            button.style.minWidth = "300px";
            button.style.minHeight = "300px";

            // 添加按钮到 body 元素中
            document.body.appendChild(button);

            // 添加按钮的点击事件处理程序
            button.addEventListener("click", async function() {
                //var if1 = document.getElementById("previewLinkOwaIframe");
                //await getIframeDocument(if1);
                //console.log(if1.contentDocument);
                //var if2 = if1.contentDocument.getElementById("preview_frame");
                //console.log(if2);
                console.log(window.location.href)
                await runJspdf();
            });
        }
    }
})();
