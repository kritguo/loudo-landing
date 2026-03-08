import{_ as Fe,C as Xe,o as i,c as R,b as t,w as a,l as P,m as A,r as v,i as f,q as Be,a as m,t as r,h as o,F as ge,s as Ae,z as ke,g as He,x as Pe,A as Ge,E as c,p as Qe,D as Ie,y as We}from"./index-Bi030A0V.js";import{o as E,a as be}from"./request-D7QyevC3.js";const Ze={class:"order-import"},Je={key:0,style:{"margin-top":"16px"}},et={class:"preview-header"},tt={key:1,class:"import-result",style:{"margin-top":"12px"}},at={key:0,style:{"margin-top":"8px","max-height":"300px","overflow-y":"auto"}},nt={key:0,style:{"margin-top":"16px"}},lt={key:1,style:{"margin-left":"4px"}},st={key:2,style:{"margin-top":"12px",display:"flex",gap:"8px"}},ot={key:1,style:{"margin-top":"16px"}},it={style:{"font-family":"monospace","font-size":"11px"}},ut={key:0,style:{"margin-top":"12px"}},rt={class:"api-section"},dt={class:"api-section-header"},pt={class:"api-key-cell"},ct={class:"api-section",style:{"margin-top":"20px"}},vt={class:"api-section",style:{"margin-top":"20px"}},ft={class:"api-section-header"},mt={key:0,style:{"margin-top":"8px","text-align":"center"}},yt={class:"log-area"},_t={__name:"OrderImport",setup(gt){const G=f("import"),D=f([]),$=f(!1),x=Ie({current:0,total:0,batchIndex:0,batchTotal:0}),j=f([]),F=f([]);function Ke(l){F.value=[],l&&l.length>0&&setTimeout(()=>{const e=l[0];F.value=[{name:e.name,raw:e}],re({raw:e})},0)}const K=f({orderNo:"",packageType:"pro_2year",vipLevel:"vip",durationDays:730,phone:"",productName:"2年PRO",sku:"PRO_2YEAR"}),Q=f(!1),S=f("import"),M=f(""),W=f(""),p=f(null),_=f(null),Z=f(!1),J=f(!1),q=f(!1),ee=f(!1),g=f({vipLevel:"vip",durationDays:730,isLifetime:!1,phone:""});function Re(){g.value={vipLevel:p.value.vipLevel||"vip",durationDays:p.value.durationDays||730,isLifetime:!!p.value.isLifetime,phone:p.value.phone||"",status:p.value.status==="canceled"?"pending":p.value.status},q.value=!0}function Le(l){l?g.value.durationDays=null:g.value.durationDays=730}async function Ve(){ee.value=!0;try{const l={orderNo:p.value.orderNo,vipLevel:g.value.vipLevel,durationDays:g.value.isLifetime?null:g.value.durationDays,productName:p.value.productName,phone:g.value.phone,status:g.value.status},e=await E("modifyOrder",l);e!=null&&e.success?(c.success("修改成功"),q.value=!1,te()):c.error((e==null?void 0:e.message)||"修改失败")}catch(l){c.error("修改失败: "+l.message)}finally{ee.value=!1}}const he={pro_2year:{vipLevel:"vip",durationDays:730,productName:"2年PRO",sku:"PRO_2YEAR"},pro_lifetime:{vipLevel:"vip",durationDays:null,productName:"永久PRO",sku:"PRO_LIFETIME"},svip_2year:{vipLevel:"svip",durationDays:730,productName:"2年MAX",sku:"MAX_2YEAR"},svip_lifetime:{vipLevel:"svip",durationDays:null,productName:"永久MAX",sku:"MAX_LIFETIME"}};function xe(){const l=he[K.value.packageType];l&&Object.assign(K.value,l)}function Ue(l){return l==="used"?"success":l==="canceled"?"danger":"warning"}function we(l){return{pending:"待激活",used:"已激活",canceled:"已取消"}[l]||l}const C=f(null),Ee=["退款","已关闭","已退款","退货","售后成功","关闭"];function Ce(l){return l?Ee.some(e=>l.includes(e)):!1}function De(l){if(!l)return null;const e=l.toUpperCase(),I=e.includes("MAX")?"svip":"vip";let s="";if(e.includes("永久")||e.includes("不限时")||e.includes("终身"))s="永久";else if(e.includes("三年")||e.includes("3年"))s="3年";else if(e.includes("两年")||e.includes("2年"))s="2年";else if(e.includes("一年")||e.includes("1年"))s="1年";else if(e.includes("半年")||e.includes("6个月"))s="半年";else return null;return{level:I,duration:s}}const T=Qe(()=>{const l=D.value.filter(e=>e._isRefund).length;return{refund:l,valid:D.value.length-l}});function Ne({row:l}){return l._isRefund?"row-refund":l._parsedLevel?"":"row-warn"}async function re(l){if(l.raw){C.value=null;try{const e=await Ge(()=>import("./xlsx-CkFp8p6R.js"),[]),I=await l.raw.arrayBuffer(),s=e.read(I,{type:"array"}),O=e.utils.sheet_to_json(s.Sheets[s.SheetNames[0]]),d={orderNo:["订单id","订单ID","订单号","订单编号","主订单号","主订单编号"],sku:["商品id","商品ID","SKU编码"],phone:["手机号","买家手机号","收货人电话"],productName:["商品名称","商品名","选购商品","物品信息"],orderStatus:["订单状态","订单售后状态","售后状态","状态"]};D.value=O.map(h=>{const u={};for(const[y,w]of Object.entries(d))for(const L of w)if(h[L]!==void 0&&h[L]!==""){u[y]=String(h[L]).trim();break}u._isRefund=Ce(u.orderStatus);const U=De(u.productName);return U&&(u._parsedLevel=U.level,u._parsedDuration=U.duration),u}).filter(h=>h.orderNo);const k=T.value;c.success(`解析到 ${D.value.length} 条订单（有效 ${k.valid}，退款 ${k.refund}）`)}catch(e){c.error("解析失败: "+e.message)}}}async function Oe(){$.value=!0,j.value=[],C.value=null;const l=200,e=D.value,I=Math.ceil(e.length/l);Object.assign(x,{current:0,total:e.length,batchIndex:0,batchTotal:I});const s={success:0,skippedDuplicate:0,skippedRefund:0,skippedRefundVipRevoked:0,skippedRefundVoided:0,canceledVipRevoked:0,canceledVoided:0,restored:0,updated:0,error:0};let O=[];try{for(let u=0;u<I;u++){x.batchIndex=u+1;const U=e.slice(u*l,(u+1)*l);try{const y=await E("importOrders",{orders:U});if(y!=null&&y.success&&y.data&&(s.success+=y.data.successCount||0,s.skippedDuplicate+=y.data.skippedDuplicate||0,s.skippedRefund+=y.data.skippedRefund||0,s.skippedRefundVipRevoked+=y.data.skippedRefundVipRevoked||0,s.skippedRefundVoided+=y.data.skippedRefundVoided||0,s.canceledVipRevoked+=y.data.canceledVipRevokedCount||0,s.canceledVoided+=y.data.canceledVoidedCount||0,s.restored+=y.data.restoredCount||0,s.updated+=y.data.updatedCount||0,s.error+=y.data.errorCount||0,y.data.details)){const w=y.data.details.filter(L=>L.type!=="duplicate"&&L.type!=="refund_skip");O=O.concat(w)}}catch(y){s.error+=U.length,j.value.push({type:"error",text:`第 ${u+1} 批导入失败: ${y.message}`})}x.current=Math.min((u+1)*l,e.length)}const d=[];s.success>0&&d.push(`新导入 ${s.success}`),s.updated>0&&d.push(`覆盖更新 ${s.updated}`),s.skippedDuplicate>0&&d.push(`已存在跳过 ${s.skippedDuplicate}`),s.skippedRefundVipRevoked>0&&d.push(`退款已激活已取消会员（之前处理）${s.skippedRefundVipRevoked}`),s.skippedRefundVoided>0&&d.push(`退款未激活已作废（之前处理）${s.skippedRefundVoided}`);const k=s.skippedRefund-s.skippedRefundVipRevoked-s.skippedRefundVoided;k>0&&d.push(`退款无记录跳过 ${k}`),s.canceledVipRevoked>0&&d.push(`退款已激活→取消会员 ${s.canceledVipRevoked}`),s.canceledVoided>0&&d.push(`退款未激活→作废 ${s.canceledVoided}`),s.restored>0&&d.push(`恢复映射 ${s.restored}`),s.error>0&&d.push(`失败 ${s.error}`);const h=`导入完成：${d.length>0?d.join("，"):"无变化"}`;C.value={...s,errorCount:s.error,details:O,message:h},j.value.push({type:"success",text:h}),c.success(h)}catch(d){j.value.push({type:"error",text:"导入失败: "+d.message})}finally{$.value=!1}}async function Ye(){if(!K.value.orderNo){c.warning("请输入订单号");return}Q.value=!0;try{const l=await E("createSingleOrder",K.value);l!=null&&l.success&&c.success("创建成功")}catch{c.error("创建失败")}finally{Q.value=!1}}async function te(){var l,e,I;if(!M.value){c.warning("请输入订单号");return}if(p.value=null,_.value=null,q.value=!1,S.value==="import")try{const s=await E("checkOrder",{orderNo:M.value,phone:W.value});s!=null&&s.success&&s.data?(p.value=s.data.selected||((l=s.data.matched)==null?void 0:l[0])||null,p.value||c.warning("未找到订单")):c.warning((s==null?void 0:s.message)||"未找到订单")}catch{c.error("查询失败")}else try{const s=await be("searchUser",{keyword:M.value,searchType:"orderId"});s!=null&&s.success&&((I=(e=s.data)==null?void 0:e.extra)!=null&&I.order)?_.value=s.data.extra.order:c.warning("未找到支付订单")}catch{c.error("查询失败")}}async function je(){var l;if((l=p.value)!=null&&l.orderNo){Z.value=!0;try{const e=await E("cancelOrder",{orderNo:p.value.orderNo});e!=null&&e.success&&(c.success("订单已取消"),p.value.status="canceled")}catch{c.error("取消失败")}finally{Z.value=!1}}}async function Se(){var l;if((l=_.value)!=null&&l.orderId){J.value=!0;try{const e=await be("refundPaymentOrder",{orderId:_.value.orderId});e!=null&&e.success&&(c.success(e.message||"已标记退款"),_.value.status="refunded",_.value.refundTime=new Date)}catch{c.error("退款失败")}finally{J.value=!1}}}const de=f([]),ae=f(!1),X=Ie({}),pe=f([]),ne=f(!1),le=f(1),se=f(0);Xe(G,l=>{l==="api"&&(B(),oe())});async function B(){ae.value=!0;try{const l=await E("manageApiKeys",{subAction:"list"});l!=null&&l.success&&(de.value=l.data||[])}catch{c.error("加载 API Key 失败")}finally{ae.value=!1}}async function qe(){const{value:l}=await We.prompt("请输入用途描述（如：抖音机器人）","创建 API Key",{confirmButtonText:"创建",cancelButtonText:"取消",inputValue:"抖音机器人"}).catch(()=>({value:null}));if(l)try{const e=await E("manageApiKeys",{subAction:"create",name:l});e!=null&&e.success&&(c.success("API Key 已创建"),B())}catch{c.error("创建失败")}}async function Te(l){try{const e=await E("manageApiKeys",{subAction:"toggle",keyId:l._id});e!=null&&e.success&&(c.success(e.message),B())}catch{c.error("操作失败")}}async function ze(l){try{const e=await E("manageApiKeys",{subAction:"delete",keyId:l._id});e!=null&&e.success&&(c.success("已删除"),B())}catch{c.error("删除失败")}}async function oe(){var l,e;ne.value=!0;try{const I=await E("getApiLogs",{page:le.value,pageSize:20});I!=null&&I.success&&(pe.value=((l=I.data)==null?void 0:l.logs)||[],se.value=((e=I.data)==null?void 0:e.total)||0)}catch{c.error("加载日志失败")}finally{ne.value=!1}}function $e(l){navigator.clipboard.writeText(l),c.success("已复制")}return(l,e)=>{const I=v("UploadFilled"),s=v("el-icon"),O=v("el-upload"),d=v("el-tag"),k=v("el-button"),h=v("el-progress"),u=v("el-table-column"),U=v("el-table"),y=v("el-alert"),w=v("el-tab-pane"),L=v("el-input"),b=v("el-form-item"),z=v("el-option"),ce=v("el-select"),N=v("el-radio"),H=v("el-radio-group"),ve=v("el-input-number"),ie=v("el-form"),V=v("el-descriptions-item"),fe=v("el-descriptions"),ue=v("el-popconfirm"),me=v("el-tabs"),Me=v("el-pagination"),ye=v("el-card"),_e=Be("loading");return i(),R("div",Ze,[t(ye,null,{header:a(()=>[...e[17]||(e[17]=[m("span",null,"订单管理",-1)])]),default:a(()=>[t(me,{modelValue:G.value,"onUpdate:modelValue":e[16]||(e[16]=n=>G.value=n)},{default:a(()=>[t(w,{label:"Excel导入",name:"import"},{default:a(()=>[t(O,{"auto-upload":!1,"on-change":re,"file-list":F.value,"onUpdate:fileList":e[0]||(e[0]=n=>F.value=n),accept:".xlsx,.xls",limit:1,"on-exceed":Ke,drag:""},{default:a(()=>[t(s,{size:40},{default:a(()=>[t(I)]),_:1}),e[18]||(e[18]=m("div",null,"拖放抖音等平台导出的订单 Excel",-1))]),_:1},8,["file-list"]),D.value.length>0?(i(),R("div",Je,[m("div",et,[m("div",null,[m("span",null,"解析到 "+r(D.value.length)+" 条订单",1),T.value.refund>0?(i(),P(d,{key:0,type:"info",size:"small",style:{"margin-left":"8px"}},{default:a(()=>[o("退款 "+r(T.value.refund),1)]),_:1})):A("",!0),T.value.valid>0?(i(),P(d,{key:1,type:"success",size:"small",style:{"margin-left":"4px"}},{default:a(()=>[o("有效 "+r(T.value.valid),1)]),_:1})):A("",!0)]),t(k,{type:"primary",loading:$.value,onClick:Oe},{default:a(()=>[o(r($.value?`导入中 ${x.batchIndex}/${x.batchTotal} 批`:`开始导入 (${T.value.valid} 条)`),1)]),_:1},8,["loading"])]),$.value&&x.total>0?(i(),P(h,{key:0,percentage:Math.round(x.current/x.total*100),format:()=>`${x.current} / ${x.total}`,style:{"margin-top":"8px"}},null,8,["percentage","format"])):A("",!0),t(U,{data:D.value.slice(0,50),size:"small","max-height":"400","row-class-name":Ne},{default:a(()=>[t(u,{prop:"orderNo",label:"订单号",width:"200"}),t(u,{label:"商品名/识别","min-width":"200"},{default:a(({row:n})=>[m("div",null,r(n.productName),1),n._parsedLevel?(i(),P(d,{key:0,size:"small",type:n._parsedLevel==="svip"?"warning":""},{default:a(()=>[o(r(n._parsedLevel==="svip"?"MAX":"PRO")+" "+r(n._parsedDuration),1)]),_:2},1032,["type"])):A("",!0)]),_:1}),t(u,{label:"状态",width:"120"},{default:a(({row:n})=>[n._isRefund?(i(),P(d,{key:0,type:"danger",size:"small"},{default:a(()=>[...e[19]||(e[19]=[o("退款/关闭",-1)])]),_:1})):n._parsedLevel?(i(),P(d,{key:2,type:"success",size:"small"},{default:a(()=>[...e[21]||(e[21]=[o("待导入",-1)])]),_:1})):(i(),P(d,{key:1,type:"warning",size:"small"},{default:a(()=>[...e[20]||(e[20]=[o("无法识别",-1)])]),_:1}))]),_:1}),t(u,{prop:"orderStatus",label:"原始状态",width:"100"})]),_:1},8,["data"]),C.value?(i(),R("div",tt,[t(y,{title:C.value.message,type:C.value.errorCount>0?"warning":"success",closable:!1},null,8,["title","type"]),C.value.details&&C.value.details.length>0?(i(),R("div",at,[(i(!0),R(ge,null,Ae(C.value.details,(n,Y)=>(i(),R("div",{key:Y,class:ke(["log-item",n.type==="error"?"error":n.type==="success"?"success":n.type==="restored"?"restored":"info"])},r(n.orderNo?`[${n.orderNo}] `:"")+r(n.message),3))),128))])):A("",!0)])):A("",!0)])):A("",!0)]),_:1}),t(w,{label:"手动创建",name:"create"},{default:a(()=>[t(ie,{model:K.value,"label-width":"100px",style:{"max-width":"500px"}},{default:a(()=>[t(b,{label:"订单号",required:""},{default:a(()=>[t(L,{modelValue:K.value.orderNo,"onUpdate:modelValue":e[1]||(e[1]=n=>K.value.orderNo=n)},null,8,["modelValue"])]),_:1}),t(b,{label:"套餐"},{default:a(()=>[t(ce,{modelValue:K.value.packageType,"onUpdate:modelValue":e[2]||(e[2]=n=>K.value.packageType=n),onChange:xe},{default:a(()=>[t(z,{label:"2年PRO",value:"pro_2year"}),t(z,{label:"永久PRO",value:"pro_lifetime"}),t(z,{label:"2年MAX",value:"svip_2year"}),t(z,{label:"永久MAX",value:"svip_lifetime"})]),_:1},8,["modelValue"])]),_:1}),t(b,{label:"VIP等级"},{default:a(()=>[t(H,{modelValue:K.value.vipLevel,"onUpdate:modelValue":e[3]||(e[3]=n=>K.value.vipLevel=n)},{default:a(()=>[t(N,{value:"vip"},{default:a(()=>[...e[22]||(e[22]=[o("PRO",-1)])]),_:1}),t(N,{value:"svip"},{default:a(()=>[...e[23]||(e[23]=[o("MAX",-1)])]),_:1})]),_:1},8,["modelValue"])]),_:1}),t(b,{label:"时长(天)"},{default:a(()=>[t(ve,{modelValue:K.value.durationDays,"onUpdate:modelValue":e[4]||(e[4]=n=>K.value.durationDays=n),min:1},null,8,["modelValue"])]),_:1}),t(b,{label:"手机号"},{default:a(()=>[t(L,{modelValue:K.value.phone,"onUpdate:modelValue":e[5]||(e[5]=n=>K.value.phone=n)},null,8,["modelValue"])]),_:1}),t(b,null,{default:a(()=>[t(k,{type:"primary",loading:Q.value,onClick:Ye},{default:a(()=>[...e[24]||(e[24]=[o("创建订单",-1)])]),_:1},8,["loading"])]),_:1})]),_:1},8,["model"])]),_:1}),t(w,{label:"订单查询",name:"check"},{default:a(()=>[t(ie,{inline:!0},{default:a(()=>[t(b,null,{default:a(()=>[t(ce,{modelValue:S.value,"onUpdate:modelValue":e[6]||(e[6]=n=>S.value=n),style:{width:"140px"}},{default:a(()=>[t(z,{label:"导入订单号",value:"import"}),t(z,{label:"微信支付单号",value:"payment"})]),_:1},8,["modelValue"])]),_:1}),t(b,null,{default:a(()=>[t(L,{modelValue:M.value,"onUpdate:modelValue":e[7]||(e[7]=n=>M.value=n),placeholder:"订单号",onKeyup:He(te,["enter"])},null,8,["modelValue"])]),_:1}),S.value==="import"?(i(),P(b,{key:0},{default:a(()=>[t(L,{modelValue:W.value,"onUpdate:modelValue":e[8]||(e[8]=n=>W.value=n),placeholder:"手机号(可选)"},null,8,["modelValue"])]),_:1})):A("",!0),t(b,null,{default:a(()=>[t(k,{type:"primary",onClick:te},{default:a(()=>[...e[25]||(e[25]=[o("查询",-1)])]),_:1})]),_:1})]),_:1}),p.value&&S.value==="import"?(i(),R("div",nt,[q.value?(i(),P(ie,{key:1,model:g.value,"label-width":"80px",style:{"max-width":"500px"}},{default:a(()=>[t(b,{label:"订单号"},{default:a(()=>[t(L,{"model-value":p.value.orderNo,disabled:""},null,8,["model-value"])]),_:1}),t(b,{label:"VIP等级"},{default:a(()=>[t(H,{modelValue:g.value.vipLevel,"onUpdate:modelValue":e[9]||(e[9]=n=>g.value.vipLevel=n)},{default:a(()=>[t(N,{value:"vip"},{default:a(()=>[...e[26]||(e[26]=[o("PRO",-1)])]),_:1}),t(N,{value:"svip"},{default:a(()=>[...e[27]||(e[27]=[o("MAX",-1)])]),_:1})]),_:1},8,["modelValue"])]),_:1}),t(b,{label:"时长"},{default:a(()=>[t(H,{modelValue:g.value.isLifetime,"onUpdate:modelValue":e[10]||(e[10]=n=>g.value.isLifetime=n),onChange:Le},{default:a(()=>[t(N,{value:!1},{default:a(()=>[...e[28]||(e[28]=[o("限时",-1)])]),_:1}),t(N,{value:!0},{default:a(()=>[...e[29]||(e[29]=[o("永久",-1)])]),_:1})]),_:1},8,["modelValue"]),g.value.isLifetime?A("",!0):(i(),P(ve,{key:0,modelValue:g.value.durationDays,"onUpdate:modelValue":e[11]||(e[11]=n=>g.value.durationDays=n),min:1,style:{"margin-left":"12px"}},null,8,["modelValue"])),g.value.isLifetime?A("",!0):(i(),R("span",lt,"天"))]),_:1}),p.value.status==="canceled"?(i(),P(b,{key:0,label:"状态"},{default:a(()=>[t(H,{modelValue:g.value.status,"onUpdate:modelValue":e[12]||(e[12]=n=>g.value.status=n)},{default:a(()=>[t(N,{value:"pending"},{default:a(()=>[...e[30]||(e[30]=[o("恢复为待激活",-1)])]),_:1}),t(N,{value:"canceled"},{default:a(()=>[...e[31]||(e[31]=[o("保持已取消",-1)])]),_:1})]),_:1},8,["modelValue"])]),_:1})):A("",!0),t(b,{label:"手机号"},{default:a(()=>[t(L,{modelValue:g.value.phone,"onUpdate:modelValue":e[13]||(e[13]=n=>g.value.phone=n),placeholder:"手机号"},null,8,["modelValue"])]),_:1}),t(b,null,{default:a(()=>[t(k,{type:"primary",loading:ee.value,onClick:Ve},{default:a(()=>[...e[32]||(e[32]=[o("保存",-1)])]),_:1},8,["loading"]),t(k,{onClick:e[14]||(e[14]=n=>q.value=!1)},{default:a(()=>[...e[33]||(e[33]=[o("取消",-1)])]),_:1})]),_:1})]),_:1},8,["model"])):(i(),P(fe,{key:0,column:2,border:""},{default:a(()=>[t(V,{label:"订单号"},{default:a(()=>[o(r(p.value.orderNo),1)]),_:1}),t(V,{label:"状态"},{default:a(()=>[t(d,{type:Ue(p.value.status),size:"small"},{default:a(()=>[o(r(we(p.value.status)),1)]),_:1},8,["type"]),p.value.activatedAt?(i(),P(d,{key:0,type:"success",size:"small",style:{"margin-left":"4px"}},{default:a(()=>[o(" 激活于 "+r(new Date(p.value.activatedAt).toLocaleString()),1)]),_:1})):A("",!0)]),_:1}),t(V,{label:"VIP等级"},{default:a(()=>[o(r(p.value.vipLevel==="svip"?"MAX":"PRO"),1)]),_:1}),t(V,{label:"时长"},{default:a(()=>[o(r(p.value.isLifetime?"永久":p.value.durationDays+"天"),1)]),_:1}),t(V,{label:"手机号"},{default:a(()=>[o(r(p.value.phone||"-"),1)]),_:1}),t(V,{label:"商品名"},{default:a(()=>[o(r(p.value.productName||"-"),1)]),_:1})]),_:1})),q.value?A("",!0):(i(),R("div",st,[t(k,{type:"primary",size:"small",onClick:Re},{default:a(()=>[...e[34]||(e[34]=[o("编辑",-1)])]),_:1}),p.value.status!=="canceled"?(i(),P(ue,{key:0,title:"确认取消此订单？已激活的将同时撤销VIP",onConfirm:je},{reference:a(()=>[t(k,{type:"danger",size:"small",loading:Z.value},{default:a(()=>[...e[35]||(e[35]=[o("取消订单(退款)",-1)])]),_:1},8,["loading"])]),_:1})):A("",!0),p.value.status==="canceled"?(i(),P(d,{key:1,type:"info"},{default:a(()=>[...e[36]||(e[36]=[o("该订单已取消",-1)])]),_:1})):A("",!0)]))])):A("",!0),_.value&&S.value==="payment"?(i(),R("div",ot,[t(fe,{column:2,border:""},{default:a(()=>[t(V,{label:"支付单号"},{default:a(()=>[o(r(_.value.orderId),1)]),_:1}),t(V,{label:"状态"},{default:a(()=>[t(d,{type:_.value.status==="paid"?"success":_.value.status==="refunded"?"danger":"warning",size:"small"},{default:a(()=>[o(r(_.value.status==="paid"?"已支付":_.value.status==="refunded"?"已退款":_.value.status),1)]),_:1},8,["type"])]),_:1}),t(V,{label:"金额"},{default:a(()=>[o("¥"+r(_.value.amount||0),1)]),_:1}),t(V,{label:"套餐"},{default:a(()=>[o(r(_.value.plan||"-"),1)]),_:1}),t(V,{label:"用户OpenID"},{default:a(()=>[m("span",it,r(_.value.openid||"-"),1)]),_:1}),t(V,{label:"支付时间"},{default:a(()=>[o(r(_.value.payTime?new Date(_.value.payTime).toLocaleString():"-"),1)]),_:1})]),_:1}),_.value.status==="paid"?(i(),R("div",ut,[t(ue,{title:"确认标记退款？将同时撤销用户VIP",onConfirm:Se},{reference:a(()=>[t(k,{type:"danger",size:"small",loading:J.value},{default:a(()=>[...e[37]||(e[37]=[o("标记退款",-1)])]),_:1},8,["loading"])]),_:1})])):A("",!0),_.value.status==="refunded"?(i(),P(d,{key:1,type:"info",style:{"margin-top":"12px"}},{default:a(()=>[o(" 退款时间: "+r(_.value.refundTime?new Date(_.value.refundTime).toLocaleString():"-"),1)]),_:1})):A("",!0)])):A("",!0)]),_:1}),t(w,{label:"API接入",name:"api"},{default:a(()=>[t(y,{type:"info",closable:!1,style:{"margin-bottom":"16px"}},{title:a(()=>[...e[38]||(e[38]=[m("span",null,"API 地址：",-1),m("code",{class:"api-url"},"https://cloudbase-4ghz65bm0b8770cd.service.tcloudbase.com/orderApiGateway",-1)])]),default:a(()=>[...e[39]||(e[39]=[m("span",{style:{"font-size":"12px",color:"#909399"}},"部署后需在 云开发控制台 → 云接入 中为 orderApiGateway 开启 HTTP 触发",-1)])]),_:1}),m("div",rt,[m("div",dt,[e[41]||(e[41]=m("h4",null,"API Key 管理",-1)),t(k,{type:"primary",size:"small",onClick:qe},{default:a(()=>[...e[40]||(e[40]=[o("创建 Key",-1)])]),_:1})]),Pe((i(),P(U,{data:de.value,size:"small"},{default:a(()=>[t(u,{label:"名称",prop:"name",width:"140"}),t(u,{label:"Key"},{default:a(({row:n})=>[m("div",pt,[m("code",null,r(X[n._id]?n.key:n.key.slice(0,12)+"..."),1),t(k,{link:"",size:"small",onClick:Y=>X[n._id]=!X[n._id]},{default:a(()=>[o(r(X[n._id]?"隐藏":"显示"),1)]),_:2},1032,["onClick"]),t(k,{link:"",size:"small",type:"primary",onClick:Y=>$e(n.key)},{default:a(()=>[...e[42]||(e[42]=[o("复制",-1)])]),_:1},8,["onClick"])])]),_:1}),t(u,{label:"状态",width:"80"},{default:a(({row:n})=>[t(d,{type:n.enabled?"success":"danger",size:"small"},{default:a(()=>[o(r(n.enabled?"启用":"禁用"),1)]),_:2},1032,["type"])]),_:1}),t(u,{label:"调用次数",prop:"requestCount",width:"85"}),t(u,{label:"最后使用",width:"140"},{default:a(({row:n})=>[o(r(n.lastUsedAt?new Date(n.lastUsedAt).toLocaleString():"-"),1)]),_:1}),t(u,{label:"操作",width:"140"},{default:a(({row:n})=>[t(k,{link:"",size:"small",type:n.enabled?"warning":"success",onClick:Y=>Te(n)},{default:a(()=>[o(r(n.enabled?"禁用":"启用"),1)]),_:2},1032,["type","onClick"]),t(ue,{title:"确认删除此 Key？",onConfirm:Y=>ze(n)},{reference:a(()=>[t(k,{link:"",size:"small",type:"danger"},{default:a(()=>[...e[43]||(e[43]=[o("删除",-1)])]),_:1})]),_:1},8,["onConfirm"])]),_:1})]),_:1},8,["data"])),[[_e,ae.value]])]),m("div",ct,[e[46]||(e[46]=m("h4",null,"调用示例",-1)),t(me,{type:"border-card",style:{"margin-top":"8px"}},{default:a(()=>[t(w,{label:"Python"},{default:a(()=>[...e[44]||(e[44]=[m("pre",{class:"code-block"},`import requests

API_URL = "https://cloudbase-4ghz65bm0b8770cd.service.tcloudbase.com/orderApiGateway"
API_KEY = "你的API Key"

# 批量导入订单
resp = requests.post(API_URL, json={
    "action": "importOrders",
    "apiKey": API_KEY,
    "orders": [
        {
            "orderNo": "抖音订单号",
            "phone": "手机号",
            "productName": "PRO永久会员",  # 自动识别等级和时长
            "amount": 99.9,
            "payTime": "2024-01-01 12:00:00",
            "platformStatus": "paid"        # paid/refunded
        }
    ]
})
print(resp.json())

# 退款同步
resp = requests.post(API_URL, json={
    "action": "updateStatus",
    "apiKey": API_KEY,
    "orderNo": "抖音订单号",
    "status": "refunded",
    "reason": "买家退款"
})

# 恢复已退款订单（退款撤销/恢复发货等）
resp = requests.post(API_URL, json={
    "action": "updateStatus",
    "apiKey": API_KEY,
    "orderNo": "抖音订单号",
    "status": "paid",              # paid/active/restored
    "reason": "退款撤销，恢复订单"
})

# 查询订单
resp = requests.post(API_URL, json={
    "action": "queryOrders",
    "apiKey": API_KEY,
    "orderNo": "抖音订单号"
})

# ========== 用户管理 API ==========

# 搜索用户（支持 nickname/openid/phone/cardNumber/orderNo）
resp = requests.post(API_URL, json={
    "action": "searchUser",
    "apiKey": API_KEY,
    "keyword": "用户昵称",
    "searchType": "nickname"
})

# 获取用户详情（含VIP、朗读数据、卡密、订单）
resp = requests.post(API_URL, json={
    "action": "getUserDetail",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# 用户列表（分页，支持 filter: all/vip/normal/expiring）
resp = requests.post(API_URL, json={
    "action": "getUserList",
    "apiKey": API_KEY,
    "page": 1,
    "pageSize": 20,
    "filter": "vip"
})

# 修改用户朗读数据
resp = requests.post(API_URL, json={
    "action": "updateUserProgress",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "progress": {
        "totalDays": 30,
        "totalReadCount": 500,
        "currentStreak": 7,
        "totalDuration": 36000,
        "lastPracticeDate": "2026-03-01"
    }
})

# 修改用户VIP状态
resp = requests.post(API_URL, json={
    "action": "updateUserVip",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "isVip": True,
    "vipLevel": "svip",
    "vipExpireTime": "2027-03-01T00:00:00",
    "isLifetime": False
})

# 给用户增加VIP天数
resp = requests.post(API_URL, json={
    "action": "addVipDays",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "days": 365,
    "vipLevel": "vip"
})

# 查询用户所有订单/卡密记录
resp = requests.post(API_URL, json={
    "action": "getUserOrders",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# ========== 卡密管理 API ==========

# 生成卡密（vipLevel: vip/svip, duration: day1/day3/month/halfYear/year/year2/year3/lifetime）
resp = requests.post(API_URL, json={
    "action": "generateCards",
    "apiKey": API_KEY,
    "count": 5,
    "vipLevel": "vip",
    "duration": "year"
})

# 查询卡密状态
resp = requests.post(API_URL, json={
    "action": "checkCard",
    "apiKey": API_KEY,
    "cardNumber": "Y12345678ABCDEF"
})

# 作废卡密（仅未使用的）
resp = requests.post(API_URL, json={
    "action": "voidCard",
    "apiKey": API_KEY,
    "cardNumber": "Y12345678ABCDEF"
})

# 激活卡密（为用户开通VIP）
resp = requests.post(API_URL, json={
    "action": "activateCard",
    "apiKey": API_KEY,
    "cardNumber": "Y12345678ABCDEF",
    "openid": "用户openid"
})

# 撤销卡密（取消用户VIP）
resp = requests.post(API_URL, json={
    "action": "revokeCard",
    "apiKey": API_KEY,
    "cardNumber": "Y12345678ABCDEF",
    "reason": "退款"
})

# ========== 智能朗读计划 API ==========

# 获取用户当前计划（含句子）
resp = requests.post(API_URL, json={
    "action": "getSmartPlan",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# 获取用户所有计划
resp = requests.post(API_URL, json={
    "action": "getUserPlans",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "status": "all"  # all/active/abandoned/completed
})

# 获取计划句子（按周/天筛选）
resp = requests.post(API_URL, json={
    "action": "getSmartSentences",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "planId": "计划ID",
    "week": 1,
    "day": 1
})

# 获取配额信息
resp = requests.post(API_URL, json={
    "action": "getQuotaInfo",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# 获取知识卡片
resp = requests.post(API_URL, json={
    "action": "getKnowledgeCards",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# ========== 句子/分类 API ==========

# 获取句子库（支持分类筛选、关键词、分页）
resp = requests.post(API_URL, json={
    "action": "getSentences",
    "apiKey": API_KEY,
    "categoryId": "framework",
    "page": 1,
    "pageSize": 20
})

# 搜索句子
resp = requests.post(API_URL, json={
    "action": "searchSentences",
    "apiKey": API_KEY,
    "keyword": "搜索关键词",
    "limit": 30
})

# 获取所有分类
resp = requests.post(API_URL, json={
    "action": "getCategories",
    "apiKey": API_KEY
})

# 获取子分类
resp = requests.post(API_URL, json={
    "action": "getSubCategories",
    "apiKey": API_KEY,
    "categoryId": "framework"
})

# ========== 统计/排行 API ==========

# 后台统计概览
resp = requests.post(API_URL, json={
    "action": "getStats",
    "apiKey": API_KEY
})

# 收入统计
resp = requests.post(API_URL, json={
    "action": "getRevenueStats",
    "apiKey": API_KEY
})

# 排行榜（type: week/year）
resp = requests.post(API_URL, json={
    "action": "getRankings",
    "apiKey": API_KEY,
    "type": "week",
    "limit": 20
})

# ========== 用户数据 API ==========

# 用户收藏
resp = requests.post(API_URL, json={
    "action": "getUserFavorites",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# 用户设置和基本信息
resp = requests.post(API_URL, json={
    "action": "getUserSettings",
    "apiKey": API_KEY,
    "openid": "用户openid"
})

# 复述记录
resp = requests.post(API_URL, json={
    "action": "getRetellingRecords",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "page": 1
})

# 练习记录
resp = requests.post(API_URL, json={
    "action": "getPracticeHistory",
    "apiKey": API_KEY,
    "openid": "用户openid",
    "limit": 50
})`,-1)])]),_:1}),t(w,{label:"curl"},{default:a(()=>[...e[45]||(e[45]=[m("pre",{class:"code-block"},`curl -X POST \\
  https://cloudbase-4ghz65bm0b8770cd.service.tcloudbase.com/orderApiGateway \\
  -H "Content-Type: application/json" \\
  -H "X-API-Key: 你的API Key" \\
  -d '{
    "action": "importOrders",
    "orders": [{
      "orderNo": "抖音订单号",
      "productName": "PRO永久会员",
      "phone": "13800138000",
      "amount": 99.9,
      "platformStatus": "paid"
    }]
  }'`,-1)])]),_:1})]),_:1})]),m("div",vt,[m("div",ft,[e[48]||(e[48]=m("h4",null,"最近调用日志",-1)),t(k,{size:"small",onClick:oe},{default:a(()=>[...e[47]||(e[47]=[o("刷新",-1)])]),_:1})]),Pe((i(),P(U,{data:pe.value,size:"small","max-height":"300"},{default:a(()=>[t(u,{label:"时间",width:"160"},{default:a(({row:n})=>[o(r(n.createdAt?new Date(n.createdAt).toLocaleString():"-"),1)]),_:1}),t(u,{label:"Action",prop:"action",width:"130"}),t(u,{label:"Key",prop:"apiKeyPrefix",width:"140"}),t(u,{label:"结果",width:"70"},{default:a(({row:n})=>[t(d,{type:n.success?"success":"danger",size:"small"},{default:a(()=>[o(r(n.success?"成功":"失败"),1)]),_:2},1032,["type"])]),_:1}),t(u,{label:"消息",prop:"message","show-overflow-tooltip":""}),t(u,{label:"耗时",width:"75"},{default:a(({row:n})=>[o(r(n.durationMs)+"ms",1)]),_:1})]),_:1},8,["data"])),[[_e,ne.value]]),se.value>20?(i(),R("div",mt,[t(Me,{"current-page":le.value,"onUpdate:currentPage":e[15]||(e[15]=n=>le.value=n),"page-size":20,total:se.value,layout:"prev, pager, next",small:"",onCurrentChange:oe},null,8,["current-page","total"])])):A("",!0)])]),_:1})]),_:1},8,["modelValue"])]),_:1}),j.value.length>0?(i(),P(ye,{key:0,style:{"margin-top":"16px"}},{header:a(()=>[...e[49]||(e[49]=[m("span",null,"导入日志",-1)])]),default:a(()=>[m("div",yt,[(i(!0),R(ge,null,Ae(j.value,(n,Y)=>(i(),R("div",{key:Y,class:ke(["log-item",n.type])},r(n.text),3))),128))])]),_:1})):A("",!0)])}}},Pt=Fe(_t,[["__scopeId","data-v-1958a24c"]]);export{Pt as default};
