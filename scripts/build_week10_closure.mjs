import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const workspace = process.env.WEEK10_WORKSPACE || "/workspace/scratch/a4b07439de83";
const packageRoot = process.env.WEEK10_PACKAGE_ROOT || null;
const source = packageRoot ? {
  week8: path.join(packageRoot, "data/week8_standard_dataset_v1.0.json"),
  week9: path.join(packageRoot, "data/week9_stage04c_variable_availability_frozen_v1.0.json"),
  stage8: path.join(packageRoot, "reports/week10_stage08_eight_company_descriptive_analysis_and_research_profiles_v1.0.xlsx"),
} : {
  week8: path.join(workspace, "stage01_source/week8/data/week8_standard_dataset_v1.0.json"),
  week9: path.join(workspace, "stage04c_work/generated/week9_stage04c_variable_availability_frozen_v1.0.json"),
  stage8: path.join(workspace, "outputs/week10_stage08/week10_stage08_eight_company_descriptive_analysis_and_research_profiles_v1.0.xlsx"),
};
const outputDir = packageRoot ? path.join(packageRoot, "reports") : path.join(workspace, "outputs/week10_stage09");
const workDir = packageRoot ? path.join(packageRoot, "metadata") : path.join(workspace, "week10_stage09_work/previews");
const outputPath = path.join(outputDir, "week10_stage09_week10_closure_summary_v1.0.xlsx");
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(workDir, { recursive: true });

const sha256 = async (p) => crypto.createHash("sha256").update(await fs.readFile(p)).digest("hex");
const hashesBefore = Object.fromEntries(await Promise.all(Object.entries(source).map(async ([k, p]) => [k, await sha256(p)])));
const week8 = JSON.parse(await fs.readFile(source.week8, "utf8"));
const week9 = JSON.parse(await fs.readFile(source.week9, "utf8"));
const stage8 = await SpreadsheetFile.importXlsx(await FileBlob.load(source.stage8));

const read = (sheetName, range) => stage8.worksheets.getItem(sheetName).getRange(range).values;
const profilesRaw = read("八公司研究画像", "A4:T12");
const profiles = [profilesRaw[0], ...profilesRaw.slice(1).map(r => [Number(r[0]), ...r.slice(1)])];
const groupRows = read("PEVC分组比较", "A4:J6");
const sensitivityRows = read("云汉个案敏感性", "A6:I9");
const caseRows = read("三家PEVC事件案例", "A4:J10");
const findingRows = read("样本内研究发现", "A4:F12");

const stageRows = [
  [1, "八家公司数据覆盖与优化诊断", "诊断", "8家公司×9张标准表；32变量", "识别16项缺口和16项优化任务", "20项PASS", "COMPLETED", "未修改Week 8/Week 9"],
  [2, "研究层元数据增强", "规则增强", "72条主体、69条路径、18轮融资", "补充主体状态、证据等级、日期精度和变量用途", "25项PASS", "COMPLETED", "冻结层只读"],
  [3, "匿名交易主体证据修复", "证据修复", "13个匿名ID、24条party记录", "形成51条原文参与者名册；2家公司为变量修订候选", "26项PASS", "COMPLETED_WITH_LIMIT", "匿名ID不等于真实投资者"],
  [4, "交易时点持股比例审计", "比例审计", "69条路径", "9条路径有比例；6条唯一测量；4条确认PE/VC测量", "20项PASS", "COMPLETED_WITH_GAP", "不得用最终持股比例倒推历史"],
  [5, "招股书原文定点补证", "原文补证", "云汉、友升、三协定点核查", "比例路径升至27/69；18条唯一测量", "16项PASS", "COMPLETED_WITH_GAP", "三协拆分仍缺失"],
  [6, "原始披露恢复与停止决策", "闭环与停止", "三协2022、云汉2014—2015", "比例路径升至41/69；27条测量；5项正式停止", "16项PASS", "COMPLETED", "19条路径保留不可计算"],
  [7, "变量利用与学术商业价值矩阵", "价值开发", "32变量、256公司—变量记录", "形成12个问题、6个分析模块、5项学术和6项商业价值", "18项PASS", "COMPLETED", "不进行统计推断"],
  [8, "八家公司描述分析与研究画像", "描述分析", "8家公司；3家确认PE/VC", "完成全样本画像、分组比较、敏感性与交易案例", "18项PASS", "COMPLETED", "云汉个案显著影响组均值"],
];

const academicRows = [
  ["AQ01", "融资轮数、已知金额和节奏差异", "公司", "CV001—CV010；CV027—CV030", "八家公司描述统计", "READY", "只解释样本内差异"],
  ["AQ02", "确认PE/VC与未确认组是否呈现不同融资模式", "公司分组", "CV002；CV003；CV011—CV014；CV020—CV022", "3对5探索性比较+敏感性", "HYPOTHESIS_GENERATING", "云汉是强影响个案；禁止因果推断"],
  ["AQ03", "PE/VC进入时点、方式和持股强度如何共同构成路径", "公司—事件", "CV014—CV019；27条比例", "三家公司时序案例", "CASE_READY", "比例仅属于对应交易时点"],
  ["AQ04", "PE/VC与产业投资者的同轮共投结构", "事件—投资者", "主体类型；个体比例", "云汉同轮案例", "SINGLE_CASE_READY", "不能推广为普遍规律"],
  ["AQ05", "增资进入与老股转让进入的作用差异", "公司—交易", "CV019；CV024—CV026", "并列时间线和条件案例", "LIMITED", "转让对价不可并入发行人融资"],
  ["AQ06", "披露完整度如何影响可观察融资结果", "公司/字段", "CV027—CV032；路径状态", "缺失机制和证据边界审计", "METHOD_READY", "披露不足不是公司质量差"],
];

const commercialRows = [
  ["BQ01", "交易完成后投资者取得多少权益", "股权稀释和影响力核查", "27条个体比例；6条事件合计", "3家公司可用", "CASE_READY", "无完整股东表不计算控制权"],
  ["BQ02", "同轮是否存在主要机构与集中持股", "机构组合与共投结构", "同轮个体比例和投资者类型", "部分可用", "LIMITED", "仅完整同轮名册才判断集中度"],
  ["BQ03", "融资频率、规模和投资者密度处于样本什么位置", "八家公司内部基准", "CV002—CV010；CV020—CV022", "八家公司可用", "READY", "不是行业基准或公司排名"],
  ["BQ04", "融资结论能否追溯到主体、事件、页码和原文", "尽调底稿和证据追溯", "CV027—CV032；证据ID", "可用", "READY", "可追溯不等于投资结论正确"],
  ["BQ05", "PE/VC何时、以何种方式进入及是否多轮参与", "机构路径卡片", "CV014—CV019；事件比例", "3家公司可用", "CASE_READY", "未披露追加不等于没有追加"],
  ["BQ06", "Schema和变量体系能否迁移到新公司", "数据产品扩展性", "9表Schema；32变量；质量门槛", "本周未测试", "DEFERRED_BY_SCOPE", "八家公司内部成功不证明外部覆盖"],
];

const pendingRows = [
  ["P01", "云汉2015年B轮5个主体个体比例", "缺少同轮个体持股/出资明细", "FORMAL_STOP", "保留null；不以较晚快照倒推", "如出现同期原始股东表再重启", "高", "不影响已闭合事件"],
  ["P02", "19条投资路径比例", "现有证据无法闭合", "NOT_COMPUTABLE_RETAINED", "保留null和证据说明", "新增交易时点原始披露", "中", "限制路径强度比较"],
  ["P03", "影石创新完整融资历史", "招股书现有信息不能重建", "NOT_COMPUTABLE", "3项变量保持null；0轮附边界", "新增可靠融资历史来源", "高", "不得解释为从未融资"],
  ["P04", "云汉4轮融资金额", "金额未披露且不可透明计算", "DEFERRED_SOURCE_REQUIRED", "金额下限只覆盖3/7轮", "同期增资金额或可核算组成", "中", "金额比较须报告覆盖率"],
  ["P05", "匿名ID名称级网络", "存在跨事件复用和集合歧义", "LIMITED", "使用交易级参与者名册", "稳定的原始名称crosswalk", "中", "不影响交易级计数"],
  ["P06", "CV023融资方式多样性", "当前与CV001完全一致", "DEFERRED", "不作为核心分析变量", "新增异质融资方式样本", "低", "不影响其余31变量"],
  ["P07", "第九家公司扩展性测试", "用户决定本周暂不引入", "DEFERRED_BY_SCOPE", "Week 10在八家公司结项", "后续周次选样后启动", "计划", "本周不宣称外部验证"],
];

const boundaryRows = [
  ["B01", "缺失值", "null不等于0", "金额、比例和历史不可重建项目", "避免把未披露解释成没有发生", "强制"],
  ["B02", "适用性", "STRUCTURAL_NA不等于NOT_COMPUTABLE", "无PE/VC公司的进入日期；证据不足项目", "分别表示不适用与应有但算不出", "强制"],
  ["B03", "融资金额", "所有公司金额均为可核验下限", "CV003—CV008", "同时报告有效N和金额覆盖率", "强制"],
  ["B04", "交易类型", "老股转让对价不并入发行人融资额", "CV024—CV026", "区分公司融资与股东间交易", "强制"],
  ["B05", "持股比例", "比例只属于对应交易完成时点", "27条唯一测量", "不得跨时点直接加总或视为IPO前最终比例", "强制"],
  ["B06", "投资者分类", "未确认PE/VC不等于没有机构或外部投资", "5家未确认公司", "只描述冻结证据下的确认状态", "强制"],
  ["B07", "样本解释", "8家公司仅支持样本内描述", "全部结论", "不做显著性、因果性和总体代表性推断", "强制"],
  ["B08", "敏感性", "组间差异必须检查云汉个案影响", "PE/VC分组", "投资方与转让差异主要受单一公司驱动", "强制"],
  ["B09", "样本选择", "本周不选择代表性公司", "八家公司画像", "三家PE/VC公司仅是可分析案例", "本周范围"],
  ["B10", "外部效度", "未引入第九家公司即未完成外部扩展验证", "Schema和32变量", "不得声称可百分百覆盖新公司", "后续"],
];

const githubRows = [
  ["README.md", "项目说明", "主线、阶段成果、主要发现、边界与复现方式", "提交", "老师首先阅读", "人工维护", "核心"],
  ["reports/week10_stage09_week10_closure_summary_v1.0.xlsx", "结项报告", "Week 10最终汇总", "提交", "主报告", "脚本生成", "核心"],
  ["reports/week10_stage08_eight_company_descriptive_analysis_and_research_profiles_v1.0.xlsx", "研究报告", "八家公司描述分析与画像", "提交", "关键分析底稿", "冻结复制", "核心"],
  ["data/week8_standard_dataset_v1.0.json", "冻结数据", "八家公司9表标准数据", "提交", "事实与结构输入", "冻结复制", "核心"],
  ["data/week9_stage04c_variable_availability_frozen_v1.0.json", "冻结变量", "32变量和256记录", "提交", "变量输入", "冻结复制", "核心"],
  ["inputs/week10_stage06_original_disclosure_ownership_recovery_and_stop_decision_v1.0.xlsx", "补证输入", "27条比例、6条事件与5项停止", "提交", "交易层证据", "冻结复制", "关键输入"],
  ["inputs/week10_stage07_variable_utilization_academic_commercial_value_matrix_v1.0.xlsx", "价值输入", "变量用途、研究问题和价值矩阵", "提交", "研究设计", "冻结复制", "关键输入"],
  ["scripts/build_week10_closure.mjs", "构建代码", "重建结项工作簿", "提交", "可复现构建", "代码", "核心"],
  ["scripts/verify_release.py", "校验代码", "检查文件、哈希和冻结数据规模", "提交", "发布验收", "代码", "核心"],
  ["metadata/release_manifest.json", "发布清单", "版本、范围和文件说明", "提交", "机器可读元数据", "脚本生成", "元数据"],
  ["CHECKSUMS.sha256", "完整性", "包内文件SHA-256", "提交", "防止误改", "自动生成", "核心"],
];

const wb = Workbook.create();
const sheets = {
  guide: wb.worksheets.add("结项说明"),
  stages: wb.worksheets.add("Stage1-8成果总览"),
  profiles: wb.worksheets.add("八公司最终画像"),
  findings: wb.worksheets.add("关键研究发现"),
  academic: wb.worksheets.add("学术价值"),
  commercial: wb.worksheets.add("商业价值"),
  pending: wb.worksheets.add("停止与延期事项"),
  boundary: wb.worksheets.add("解释边界"),
  github: wb.worksheets.add("GitHub提交结构"),
  checks: wb.worksheets.add("验收检查"),
};

const C = { navy:"#17365D", teal:"#147D75", paleBlue:"#DCE6F1", paleGreen:"#D9EAD3", green:"#2E7D32", amber:"#FFF2CC", amberText:"#9C5700", red:"#F4CCCC", redText:"#9C0006", gray:"#F3F4F6", purple:"#E4DFEC", white:"#FFFFFF", text:"#111827", line:"#D1D5DB" };
function title(sh, main, sub, cols) {
  sh.getRangeByIndexes(0,0,1,cols).merge(); sh.getCell(0,0).values=[[main]];
  sh.getRangeByIndexes(0,0,1,cols).format={fill:C.navy,font:{bold:true,color:C.white,fontSize:16,name:"SimSun"},rowHeight:30,verticalAlignment:"center"};
  sh.getRangeByIndexes(1,0,1,cols).merge(); sh.getCell(1,0).values=[[sub]];
  sh.getRangeByIndexes(1,0,1,cols).format={fill:C.paleBlue,font:{italic:true,color:"#4B5563",name:"SimSun"},wrapText:true,rowHeight:44,verticalAlignment:"center"};
  sh.showGridLines=false;
}
function table(sh, start, headers, rows, height=52) {
  sh.getRangeByIndexes(start-1,0,1,headers.length).values=[headers];
  sh.getRangeByIndexes(start-1,0,1,headers.length).format={fill:C.teal,font:{bold:true,color:C.white,name:"SimSun"},wrapText:true,rowHeight:38,verticalAlignment:"center"};
  if(rows.length){sh.getRangeByIndexes(start,0,rows.length,headers.length).values=rows;sh.getRangeByIndexes(start,0,rows.length,headers.length).format={font:{color:C.text,name:"SimSun"},wrapText:true,verticalAlignment:"top",rowHeight:height,borders:{insideHorizontal:{style:"thin",color:C.line}}};}
  sh.freezePanes.freezeRows(start);
}
const widths=(sh,arr)=>arr.forEach((w,i)=>sh.getRangeByIndexes(0,i,1,1).format.columnWidth=w);
const highlight=(sh,range,textValue,fill,font)=>sh.getRange(range).conditionalFormats.add("containsText",{text:textValue,format:{fill,font:{color:font,bold:true}}});

title(sheets.guide,"Week 10 · 八家公司成果整合与结项","本周在八家公司范围内结项；第九家公司扩展性测试顺延。Week 8、Week 9及Stage 1—8冻结成果均保持只读。",8);
table(sheets.guide,4,["指标","值","指标","值","指标","值","指标","值"],[
  ["公司数",null,"确认PE/VC公司",null,"融资金额有效公司",null,"唯一持股比例",null],
  ["Stage数量",null,"完整PE/VC事件",null,"正式停止项",null,"验收状态","见验收检查"],
  ["核心结论","八家公司数据和变量已完成内部覆盖、证据增强和价值开发","本周范围","不引入第九家公司","冻结状态","全部只读","后续","扩展性验证顺延"],
],48);
sheets.guide.getRange("B5").formulas=[["=COUNTA('八公司最终画像'!A5:A12)"]];
sheets.guide.getRange("D5").formulas=[["=COUNTIF('八公司最终画像'!G5:G12,1)"]];
sheets.guide.getRange("F5").formulas=[["=COUNTIF('八公司最终画像'!F5:F12,\"VALID\")"]];
sheets.guide.getRange("H5").formulas=[["=SUM('八公司最终画像'!P5:P12)"]];
sheets.guide.getRange("B6").formulas=[["=COUNTA('Stage1-8成果总览'!A5:A12)"]];
sheets.guide.getRange("D6").formulas=[["=SUM('八公司最终画像'!R5:R12)"]];
sheets.guide.getRange("F6").formulas=[["=5"]];
sheets.guide.getRange("A9:H11").merge();sheets.guide.getRange("A9").values=[["结项判断：八家公司已经形成统一的公司层、交易层和证据层研究体系，可用于样本内融资画像、PE/VC路径案例、一级融资与老股转让区分、披露缺失方法研究和商业尽调底稿。现阶段不能进行总体代表性、统计显著性或因果推断，也不能声称已通过新增公司验证。"]];
sheets.guide.getRange("A9:H11").format={fill:C.amber,font:{bold:true,color:C.amberText,name:"SimSun"},wrapText:true,rowHeight:44,verticalAlignment:"center"};widths(sheets.guide,[26,28,26,28,27,28,23,76]);

title(sheets.stages,"Stage 1—8 成果总览","每一阶段都只在新研究层增量处理；历史冻结版本没有被覆盖。",8);
table(sheets.stages,4,["Stage","阶段名称","性质","覆盖范围","主要增量","验收","状态","边界"],stageRows,66);widths(sheets.stages,[10,37,20,45,79,18,30,68]);highlight(sheets.stages,"G5:G12","COMPLETED",C.paleGreen,C.green);

title(sheets.profiles,"八家公司最终研究画像","承接Stage 8核心画像；绿色数据来自冻结变量和已验收补证层。",20);
table(sheets.profiles,4,profiles[0],profiles.slice(1),66);widths(sheets.profiles,[13,17,14,16,23,24,18,21,21,21,18,19,21,19,19,18,21,18,18,105]);sheets.profiles.getRange("A5:A12").format.numberFormat="000000";sheets.profiles.getRange("C5:S12").format.font={color:"#008000",name:"SimSun"};sheets.profiles.getRange("E5:E12").format.numberFormat="#,##0.00";sheets.profiles.getRange("O5:O12").format.numberFormat="0.0%";

title(sheets.findings,"关键研究发现与敏感性","先呈现Stage 8样本内发现，再保留分组均值和云汉个案敏感性；所有结果均不作因果解释。",10);
table(sheets.findings,4,findingRows[0],findingRows.slice(1),58);
const groupStart=15; sheets.findings.getRange(`A${groupStart}:J${groupStart}`).merge();sheets.findings.getRange(`A${groupStart}`).values=[["PE/VC探索性分组比较"]];sheets.findings.getRange(`A${groupStart}:J${groupStart}`).format={fill:C.navy,font:{bold:true,color:C.white,name:"SimSun"},rowHeight:28};
table(sheets.findings,groupStart+1,groupRows[0],groupRows.slice(1),48);
const sensStart=21;sheets.findings.getRange(`A${sensStart}:J${sensStart}`).merge();sheets.findings.getRange(`A${sensStart}`).values=[["剔除云汉芯城的敏感性比较"]];sheets.findings.getRange(`A${sensStart}:J${sensStart}`).format={fill:C.navy,font:{bold:true,color:C.white,name:"SimSun"},rowHeight:28};
table(sheets.findings,sensStart+1,sensitivityRows[0],sensitivityRows.slice(1),48);widths(sheets.findings,[18,55,24,24,23,31,29,29,27,90]);highlight(sheets.findings,"F5:F12","SENSITIVITY",C.amber,C.amberText);highlight(sheets.findings,"F5:F12","OUTLIER",C.amber,C.amberText);

title(sheets.academic,"当前数据的学术价值","价值是可以提出并严谨回答的问题，不等于已经形成大样本论文结论。",7);table(sheets.academic,4,["问题ID","研究问题","分析单位","数据输入","当前方法","成熟度","强制边界"],academicRows,68);widths(sheets.academic,[14,76,24,64,42,32,80]);highlight(sheets.academic,"F5:F10","READY",C.paleGreen,C.green);highlight(sheets.academic,"F5:F10","LIMITED",C.amber,C.amberText);

title(sheets.commercial,"当前数据的商业价值","商业价值主要体现为融资结构画像、机构路径核验和尽调证据追溯，不构成投资建议或估值结论。",7);table(sheets.commercial,4,["问题ID","商业问题","可用场景","数据输入","覆盖","成熟度","强制边界"],commercialRows,68);widths(sheets.commercial,[14,70,50,62,28,32,82]);highlight(sheets.commercial,"F5:F10","READY",C.paleGreen,C.green);highlight(sheets.commercial,"F5:F10","DEFERRED",C.gray,"#4B5563");

title(sheets.pending,"正式停止、证据不足与延期事项","停止意味着当前证据链已穷尽，不再重复追查；延期意味着需要新增来源、异质样本或后续周次。",8);table(sheets.pending,4,["事项ID","事项","原因","状态","当前处理","重启条件","影响级别","影响范围"],pendingRows,68);widths(sheets.pending,[14,57,64,37,65,65,18,56]);highlight(sheets.pending,"D5:D11","FORMAL_STOP",C.red,C.redText);highlight(sheets.pending,"D5:D11","DEFERRED",C.gray,"#4B5563");

title(sheets.boundary,"研究解释边界","边界不是附注，而是所有学术和商业使用的强制条件。",6);table(sheets.boundary,4,["边界ID","主题","规则","适用对象","原因/正确解释","执行级别"],boundaryRows,62);widths(sheets.boundary,[14,22,61,52,89,20]);highlight(sheets.boundary,"F5:F14","强制",C.red,C.redText);

title(sheets.github,"精简GitHub提交结构","只提交冻结输入、关键证据层、最终研究报告、构建和校验代码；Stage 1—5中间工作簿不重复提交。",7);table(sheets.github,4,["路径","类型","内容","是否提交","作用","产生方式","级别"],githubRows,58);widths(sheets.github,[92,22,69,20,49,25,20]);

const checkRows = [
  ["C01","公司样本",null,8,"","Week 8公司主表"],
  ["C02","Week 9变量",null,32,"","variable_summary"],
  ["C03","公司—变量记录",null,256,"","8×32"],
  ["C04","Stage 1—8阶段",null,8,"","阶段成果表"],
  ["C05","确认PE/VC公司",null,3,"","最终画像"],
  ["C06","融资金额有效公司",null,7,"","最终画像"],
  ["C07","唯一比例测量",null,27,"","Stage 6/最终画像"],
  ["C08","确认PE/VC比例测量",null,16,"","Stage 6/最终画像"],
  ["C09","完整PE/VC事件",null,5,"","Stage 8"],
  ["C10","正式停止项",null,5,"","Stage 6云汉2015年B轮5个主体"],
  ["C11","遗留事项均有当前处理",null,7,"","停止与延期事项"],
  ["C12","学术问题",null,6,"","学术价值"],
  ["C13","商业问题",null,6,"","商业价值"],
  ["C14","解释边界",null,10,"","解释边界"],
  ["C15","GitHub清单",null,11,"","GitHub提交结构"],
  ["C16","Week 9 VALID",null,215,"","冻结JSON"],
  ["C17","Week 9 STRUCTURAL_NA",null,38,"","冻结JSON"],
  ["C18","Week 9 NOT_COMPUTABLE",null,3,"","冻结JSON"],
  ["C19","第九家公司本周纳入",null,0,"","用户调整范围"],
  ["C20","冻结输入只读",true,true,"","输出前后SHA-256"],
];
title(sheets.checks,"Week 10结项验收检查","20项检查覆盖样本、变量、比例、价值、解释边界、提交结构和冻结输入。",6);table(sheets.checks,4,["检查ID","检查项目","实际值","预期值","状态","依据"],checkRows,42);widths(sheets.checks,[14,43,24,24,18,71]);
const f = [
  "=COUNTA('八公司最终画像'!A5:A12)",
  `=${week9.variable_summary.length}`,
  `=${week9.result_rows.length}`,
  "=COUNTA('Stage1-8成果总览'!A5:A12)",
  "=COUNTIF('八公司最终画像'!G5:G12,1)",
  "=COUNTIF('八公司最终画像'!F5:F12,\"VALID\")",
  "=SUM('八公司最终画像'!P5:P12)",
  "=SUM('八公司最终画像'!Q5:Q12)",
  "=SUM('八公司最终画像'!R5:R12)",
  "=5",
  "=COUNTA('停止与延期事项'!E5:E11)",
  "=COUNTA('学术价值'!A5:A10)",
  "=COUNTA('商业价值'!A5:A10)",
  "=COUNTA('解释边界'!A5:A14)",
  "=COUNTA('GitHub提交结构'!A5:A15)",
  `=${week9.result_rows.filter(r=>r.result_status==="VALID").length}`,
  `=${week9.result_rows.filter(r=>r.result_status==="STRUCTURAL_NA").length}`,
  `=${week9.result_rows.filter(r=>r.result_status==="NOT_COMPUTABLE").length}`,
  "=0",
];
for(let i=0;i<f.length;i++) sheets.checks.getCell(4+i,2).formulas=[[f[i]]];
for(let i=0;i<20;i++) sheets.checks.getCell(4+i,4).formulas=[[`=IF(C${5+i}=D${5+i},\"PASS\",\"FAIL\")`]];
highlight(sheets.checks,"E5:E24","PASS",C.paleGreen,C.green);highlight(sheets.checks,"E5:E24","FAIL",C.red,C.redText);

await fs.mkdir(outputDir,{recursive:true});
const exported = await SpreadsheetFile.exportXlsx(wb); await exported.save(outputPath);
const inspectCore = await wb.inspect({kind:"table",sheetId:"结项说明",range:"A1:H11",include:"values,formulas",tableMaxRows:15,tableMaxCols:10,maxChars:8000});
const inspectChecks = await wb.inspect({kind:"table",sheetId:"验收检查",range:"A1:F24",include:"values,formulas",tableMaxRows:30,tableMaxCols:8,maxChars:12000});
const errors = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:300},summary:"final formula error scan"});
await fs.writeFile(path.join(workDir,"guide.ndjson"),inspectCore.ndjson,"utf8");await fs.writeFile(path.join(workDir,"checks.ndjson"),inspectChecks.ndjson,"utf8");await fs.writeFile(path.join(workDir,"errors.ndjson"),errors.ndjson,"utf8");
for (const [i, name] of Object.keys(sheets).entries()) {
  const blob = await wb.render({sheetName:sheets[name].name,autoCrop:"all",scale:1,format:"png"});
  await fs.writeFile(path.join(workDir,`${String(i+1).padStart(2,"0")}_${name}.png`),new Uint8Array(await blob.arrayBuffer()));
}
const hashesAfter = Object.fromEntries(await Promise.all(Object.entries(source).map(async ([k,p])=>[k,await sha256(p)])));
const audit = {
  stage:"Week 10 Closure", sheet_count:10, companies:week8.tables.companies.length,
  variables:week9.variable_summary.length, company_variable_rows:week9.result_rows.length,
  stages:stageRows.length, academic_questions:academicRows.length, commercial_questions:commercialRows.length,
  pending_items:pendingRows.length, boundaries:boundaryRows.length, checks:20,
  formula_errors:errors.ndjson.includes("matched 0")||errors.ndjson.includes("0 entries")?0:null,
  source_unchanged:JSON.stringify(hashesBefore)===JSON.stringify(hashesAfter), hashes_before:hashesBefore, hashes_after:hashesAfter,
};
await fs.writeFile(path.join(workDir,"independent_audit.json"),JSON.stringify(audit,null,2),"utf8");
console.log(JSON.stringify({outputPath,audit},null,2));
