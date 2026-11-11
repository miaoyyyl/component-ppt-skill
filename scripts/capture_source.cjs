"use strict";
// 公开来源网页的实际浏览器截图；不修改页面，不借截图伪造科研图。
const fs=require("node:fs"),dns=require("node:dns/promises"),net=require("node:net");
const {chromium}=require(process.argv[2]),request=JSON.parse(fs.readFileSync(process.argv[4],"utf8"));
function isPublic(address){
  if(net.isIPv4(address)){const [a,b]=address.split('.').map(Number);return !(a===0||a===10||a===127||a>=224||a===169&&b===254||a===172&&b>=16&&b<=31||a===192&&b===168||a===100&&b>=64&&b<=127||a===198&&(b===18||b===19));}
  const ip=address.toLowerCase();return !(ip==='::'||ip==='::1'||ip.startsWith('fc')||ip.startsWith('fd')||ip.startsWith('fe8')||ip.startsWith('fe9')||ip.startsWith('fea')||ip.startsWith('feb')||ip.startsWith('ff')||ip.startsWith('::ffff:'));
}
(async()=>{
 const browser=await chromium.launch({headless:true,executablePath:process.argv[3]});
 try{
  const context=await browser.newContext({viewport:request.viewport}),cache=new Map();
  await context.route('**/*',async route=>{
   try{
    const r=route.request(),u=new URL(r.url());
    if(!['http:','https:'].includes(u.protocol)||!['GET','HEAD'].includes(r.method())||u.username||u.password||u.port&&!['80','443'].includes(u.port))return route.abort();
    if(!cache.has(u.hostname))cache.set(u.hostname,(await dns.lookup(u.hostname,{all:true})).every(a=>isPublic(a.address)));
    return cache.get(u.hostname)?route.continue():route.abort();
   }catch{return route.abort();}
  });
  const page=await context.newPage(),response=await page.goto(request.url,{waitUntil:'domcontentloaded',timeout:45000});
  if(!response||response.status()>=400)throw new Error('来源网页无法公开读取');
  await page.locator('body').waitFor({state:'visible',timeout:5000});
  await page.evaluate(()=>Promise.race([document.fonts.ready,new Promise(resolve=>setTimeout(resolve,3000))]));
  if((await page.locator('body').innerText()).trim().length<8)throw new Error('来源页面没有可读取内容');
  await page.screenshot({path:request.output,fullPage:false});
  console.log(JSON.stringify({url:page.url(),screenshot:request.output}));
 }finally{await browser.close();}
})().catch(error=>{console.error(error.message);process.exitCode=1;});
