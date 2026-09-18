from __future__ import annotations

import json
from pathlib import Path
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
DOCX_PATH = ROOT / "《把答案写在大地上》整体音乐设计方案.docx"
JSON_PATH = ROOT / "《把答案写在大地上》镜头音乐数据.json"
README_PATH = ROOT / "README_使用说明.md"
ATTR_PATH = ROOT / "版权署名模板.txt"


def tc(seconds: int) -> str:
    return f"{seconds // 60:02d}:{seconds % 60:02d}"


shot_content = [
    ("黑场淡入空教室，窗帘轻动，片头留白", "冷静、空旷、尚未提出问题", "无人物，观众进入空间", "建立现实主义基调和留白", "否", "无配乐，仅环境声", "克制、观察", "无", "无", "静音；保留风与教室底噪", "让观众先相信空间，避免片头被音乐预设情绪", "SIL", "剧组实录风声、教室room tone"),
    ("陈老师写完问号并转身", "理性、专注", "教师沉静、准备提问", "把抽象命题变成可见动作", "否", "无配乐，突出粉笔声", "理性、悬而未决", "无", "无", "静音", "粉笔摩擦成为第一处声音焦点", "SIL", "粉笔、衣料、轻微室内混响"),
    ("教师提问，林悦停笔", "克制的疑问感", "教师引导；林悦由记录转为思考", "抛出全片核心问题", "否", "对白优先，句末留0.5秒静默", "疑问、停顿", "无", "无", "静音", "用真正的停顿制造悬念，不以音乐替代表演", "SIL", "对白、笔尖停止、呼吸、教室底噪"),
    ("笔尖悬停，镜头上仰至林悦抬眼", "内向、专注、轻微困惑", "林悦开始主动思考", "主角认知旅程启动", "是", "M01主旋律首次出现，仅毡音钢琴四音动机", "疑问、孤独、克制", "00:34.5", "00:40.0续接", "极低，音乐总线约-30 dB", "第一次把“问号主题”与林悦绑定", "M01", "笔帽轻触、纸张微响"),
    ("陈老师合书并鼓励去生活里寻找", "温和、理性", "教师给予方向，不灌输答案", "推动主角离开课堂", "是", "M01增加一层极轻空气垫，不加节拍", "被鼓励、仍保留疑问", "00:40.0", "00:50.0续接", "对白下-30 dB，句尾抬至-27 dB", "托住教师语气，避免煽情", "M01", "合书、衣料、对白"),
    ("林悦合上红色笔记本，望向窗外", "疏离转为行动意愿", "林悦决定寻找答案", "完成课堂段落的小转折", "是", "M01钢琴动机完整一次，加入单音大提琴泛音", "思索、微弱决心", "00:50.0", "01:00.0续接", "低，-27至-25 dB；画外音下再降4 dB", "将问号从静态疑问转为行动冲动", "M01", "合本、轻吸气、窗外树叶"),
    ("收起笔记本并拿出手机稳定器", "务实、准备出发", "林悦从思考进入执行", "连接课堂与社会实践", "是", "M01尾奏与M02脉冲交叉，最后2秒让脚步接管", "清醒、行动", "01:00.0", "01:08.5", "低，-26 dB渐弱至-32 dB", "用音色交接而非鼓点硬转场", "M01→M02", "包布摩擦、桌椅轻响、稳定器拿取"),
    ("林悦进入空走廊并向远端行走", "开阔、独行、方向明确", "林悦主动出发", "正式进入社会实践段", "是", "M02毡音钢琴分解音型与低强度木质脉冲", "探索、轻微期待", "01:08.0", "01:20.0续接", "低，-28至-25 dB", "给步伐一个稳定但不商业化的节奏", "M02", "走廊脚步、关门、风声"),
    ("校园步道侧移，林悦检查手机画面", "自然、流动", "专注、带着任务感", "交代距离与行动过程", "是", "M02加入极轻拨弦/木片节拍，约72 BPM", "前行、好奇", "01:20.0", "01:30.0续接", "低，-25 dB；画外音下-29 dB", "维持行进感并给自行车铃留空间", "M02", "树叶、脚步、远处单次自行车铃"),
    ("进入红砖社区窄巷，抬眼观察旧墙与木门", "陌生、安静、时间感", "林悦放慢速度并观察", "从校园进入生活现场", "是", "M02减少脉冲，保留钢琴与低频空气，和声转为开放五度", "探索、历史感初现", "01:30.0", "01:40.0续接", "低，-27 dB", "让空间质感接替行进节奏", "M02", "巷道脚步、远处生活声、轻风"),
    ("林悦与周建国在木门前初见", "克制、生活化、礼貌", "双方试探性建立信任", "引入讲述者", "是", "M02仅保留长音与两次稀疏钢琴，避开对白", "温和、真实", "01:40.0", "01:50.0续接", "对白下极低-32 dB，句间不抬升", "维持跨场连续但不抢人物第一次交流", "M02", "对白、木门触碰、巷道底噪"),
    ("周建国开门让路，林悦跨入史小屋", "从外到内、私密感增强", "林悦接受邀请，周建国开始引导", "完成空间和叙事层级转换", "是", "M02尾部过滤变暗，挂钟滴答成为下一段拍点", "沉静、进入记忆", "01:50.0", "01:59.8", "-29 dB渐弱至无", "以门轴和滴答替代音乐完成转场", "M02", "门轴、脚步、室内混响、挂钟渐入"),
    ("桌上相册、杯子、钥匙被整理到中央", "温暖、陈旧、可触摸", "周建国准备讲述；林悦专注", "建立记忆物证", "是", "M03低声弦乐与毡音钢琴进入，玻璃马林巴只点一次", "怀旧、克制、非悲情", "02:00.4", "02:10.0续接", "极低-30 dB", "让物件拥有重量，不把老人塑造成苦情人物", "M03", "杯底、相册与木桌摩擦、挂钟"),
    ("翻开相册，泥路与低矮砖房照片出现", "触觉化的过去", "周建国从物件进入回忆", "把宏大历史落到生活细节", "是", "M03大提琴泛音轻进，钢琴不落重拍", "怀旧、艰辛但平静", "02:10.0", "02:20.0续接", "旁白下-31 dB；翻页瞬间音乐让位2 dB", "强调照片与口述而不是渲染苦难", "M03", "翻页、指尖摩擦、挂钟、旁白"),
    ("周建国看照片后望向窗外，讲煤油灯", "安静、回望、略带苍凉", "老人回忆但不悲情", "深化亲历者可信度", "是", "M03增加一条低音大提琴长音，末尾留呼吸", "怀旧、微苦、温暖", "02:20.0", "02:30.0续接", "对白下-31 dB，末尾-27 dB", "承接记忆，同时保持人物表演在前", "M03", "对白、木窗轻响、挂钟"),
    ("林悦用稳定器拍摄相册并按下录制", "观察、记录", "林悦仍主要通过设备理解", "展示她的初始观看方式", "是", "M03钢琴四音动机以变奏出现，节奏保持稀疏", "专注、追问", "02:30.0", "02:40.0续接", "低-28 dB；画外音下-31 dB", "把林悦主题与老人记忆第一次连接", "M03", "稳定器电机、录制轻触、旁白"),
    ("旧钥匙被放到林悦面前", "凝视、象征性增强", "周建国珍视；林悦被吸引", "引入全片核心物件母题", "是", "M03在钥匙落桌时加入低频木质点音，随后静一拍", "重量、时间、珍惜", "02:40.0", "02:50.0续接", "对白下-31 dB；落桌前后瞬降3 dB", "让钥匙的实体声成为主题触发器", "M03", "金属钥匙轻触木桌、红绳摩擦、对白"),
    ("林悦写下路、灯、房、学堂", "理解开始形成但尚未完成", "林悦把口述整理成线索", "第一次形成可见答案框架", "是", "M03和声略转亮，钢琴动机补全最后一音", "触动、梳理", "02:50.0", "03:00.0续接", "低-27 dB；画外音下-30 dB", "把抽象历史转为四个生活关键词", "M03", "笔尖、纸张、挂钟、画外音"),
    ("二人背向镜头走出木屋，去院子看现在", "门槛、过渡、期待", "周建国引导；林悦跟随", "从历史物证转向现实空间", "是", "M03尾奏拉开高频，门外环境声增大；末1秒预示M04弦乐", "过去通向现在", "03:00.0", "03:09.7", "-27 dB渐弱，出门后-30 dB", "通过声场变宽完成时间和空间转换", "M03→M04", "合本、脚步、门口空气转换、对白"),
    ("二人进入改造后的庭院，设施全貌展开", "开阔、暖金、生活改善可见", "两人并行观察", "展示现实答案的全景", "是", "M04柔和弦乐铺开，钢琴主题转大调色彩", "温暖、开阔、克制希望", "03:10.0", "03:20.0续接", "低至中低-25 dB", "画面首次打开时打开和声，而非突然煽情", "M04", "树叶、远处生活声、脚步"),
    ("周建国沿无障碍坡道走三步并触扶手", "具体、可靠、有人情温度", "老人以身体经验验证改善", "用动作证明民生变化", "是", "M04降低弦乐密度，保留中提琴与轻钢琴", "朴实、安心", "03:20.0", "03:30.0续接", "对白下-30 dB", "让脚步和扶手触碰成为证据", "M04", "坡道脚步、手掌擦过金属扶手、对白"),
    ("林悦放低设备，改用眼睛观察电梯", "由记录转为真正看见", "林悦发生认知方式变化", "主角内在转折前置", "是", "M04撤去脉冲，四音主题由单簧管/钢琴轻奏", "领悟、专注", "03:30.0", "03:40.0续接", "画外音下-29 dB，句尾抬至-24 dB", "用减法突出“放下设备”的动作", "M04", "设备下降、衣料、庭院空气、画外音"),
    ("横移呈现坡道、电梯、阅读室和长椅", "平实、丰富、可居住", "无人物，空间本身说话", "集中呈现民生改善", "是", "M04完整展开一次主题，弦乐最饱满但无铜管", "温暖、踏实、希望", "03:40.0", "03:50.0续接", "中低-22 dB；旁白下-27 dB", "在无人物镜头承担信息连接和情绪上扬", "M04", "树叶、鸟鸣一次、远处生活声、旁白"),
    ("旧红砖楼与新设施同框，两人停下", "历史与现实并置", "老人笃定；林悦理解加深", "明确变化来自持续行动", "是", "M04回收至钢琴与大提琴，句尾留1秒空白", "稳重、接续、非口号化", "03:50.0", "03:59.6", "对白下-30 dB；最后1秒-26 dB", "为钥匙交接预留情绪空间", "M04", "对白、树叶、远处低频城市底噪"),
    ("旧钥匙在两代人的手中完成交接", "庄重、亲密、象征性强", "老人托付；林悦郑重承接", "全片核心象征动作", "是", "M05从钥匙金属声后0.3秒进入，钢琴主题低八度回答", "传承、信任、重量", "04:00.3", "04:10.0续接", "低-26 dB，手停稳后升至-23 dB", "将钥匙母题与代际连接绑定", "M05", "钥匙与红绳、手部衣料、呼吸"),
    ("林悦看钥匙再看楼门，表情由疑惑转为理解", "暖金、内在转折", "林悦真正理解生活证据", "完成主角认知转折", "是", "M05加入独奏大提琴与柔弦，主题由小调转相对大调", "触动、清晰、克制喜悦", "04:10.0", "04:20.0续接", "画外音下-28 dB，末尾-22 dB", "承载全片最重要的无对白内心变化", "M05", "画外音、庭院风、极轻金属触感"),
    ("林悦在笔记本写下四史关键词", "整理、贯通", "林悦把生活经验与理论连接", "完成理论回扣", "是", "M05主题加入高音弦乐但保持中弱，笔尖处避让", "清晰、上扬", "04:20.0", "04:30.0续接", "中低-22 dB，笔尖瞬间-25 dB", "把前面的生活关键词提升为历史理解", "M05", "笔尖、纸张、树叶"),
    ("林悦提问，周建国回答历史在脚下和手里", "平等、真诚、思想落点", "学生主动追问；老人笃定回应", "给出核心命题的生活化表达", "是", "M05几乎抽空，仅留大提琴长音；回答后主题上行两音", "笃定、温暖", "04:30.0", "04:40.0续接", "对白下-33 dB，回答结束后-24 dB", "让对白成为高潮内容，音乐只提供余韵", "M05", "对白、长椅轻响、庭院底噪"),
    ("林悦架好自拍机位并后退到标记位", "准备表达、行动明确", "林悦从观察者变为讲述者", "为自述建立节奏和期待", "是", "M05恢复轻脉冲，弦乐持续上行但不爆发", "准备、坚定", "04:40.0", "04:50.0续接", "中低-23 dB", "积蓄力量，把高潮留给自述与灯亮", "M05", "稳定器落地/支架、两步脚步、衣摆"),
    ("林悦正对镜头自述，背景路灯末尾亮起", "坚定、清澈、暖金转暮色", "林悦形成自己的答案", "全片表达高潮", "是", "M05前1秒收窄；对白下低铺，最后一句末尾和路灯同时给出和声顶点", "坚定、希望、兑现", "04:50.0", "04:59.6", "对白下-29 dB；末句后升至-19 dB后快速回落", "让高潮由人物语言和灯光共同完成，避免盖住台词", "M05", "对白、路灯继电器极轻一声、庭院风"),
    ("蓝调暮色中林悦沿走廊返回教室", "回归、目标明确", "林悦带着答案返回", "开启首尾呼应", "是", "M06回奏M01四音主题，调性明亮一阶，钢琴+低弦", "笃定、回归", "05:00.4", "05:10.0续接", "低-26 dB；画外音下-29 dB", "让离开与返回在听觉上形成闭环", "M06", "走廊脚步、门把、画外音"),
    ("林悦连接手机与投影，陈老师抬眼", "安静、期待", "学生主动展示；教师等待", "把社会实践带回课堂", "是", "M06压到单音钢琴，并在05:18开始淡出", "准备、专注", "05:10.0", "05:19.2", "极低-31 dB", "把连接提示音和投影风扇留在前景", "M06", "设备放置、插线、连接提示音、投影风扇"),
    ("教师肩后观看投影中的相册、钥匙、坡道与自述", "观看、回响、媒介内外连接", "教师专注接收学生成果", "让前序素材回到课堂", "否（外部配乐）", "只保留投影片内的M05尾音，禁止叠加M06", "回响、真实播放感", "05:20.0（片内声）", "05:29.5", "投影片声音约-28 dB并做小扬声器带宽；外部音乐静音", "避免“画外配乐+片内配乐”双层冲突", "DIEGETIC", "投影风扇、片内对白与片内音乐、教室room tone"),
    ("陈老师看向林悦并轻点头，给出认可", "克制、温暖、成熟", "教师确认学生找到方法", "完成师生反馈", "是", "M06在教师转头后0.4秒再入，钢琴主题由弦乐回答", "认可、释然", "05:30.4", "05:40.0续接", "对白下-30 dB，点头后-24 dB", "不提前提示认可，让表演先发生", "M06", "对白、椅子细响、投影风扇"),
    ("林悦提出走出去、带回来的课堂期待", "真诚、平等、面向未来", "林悦从答题者变为行动倡议者", "价值表达与行动落点", "是", "M06弦乐缓慢上扬，钢琴保持主旋律，不加鼓和铜管", "坚定、开放、希望", "05:40.0", "05:50.0续接", "对白下-29 dB，句尾-22 dB", "支持语言的行动感而不变成口号音乐", "M06", "对白、投影风扇、教室呼吸"),
    ("问号被一笔延伸为箭头，合本并渐黑", "简洁、完成、继续向前", "林悦形成方法并准备行动", "视觉与听觉收束", "是", "M06四音主题最后一次完整出现，箭头落笔对应一记柔和高音，合本后只留尾音", "释然、笃定、余韵", "05:50.0", "05:58.5；05:58.5-06:00只留room tone后黑场", "中低-21 dB渐弱至无；合本声保持清晰", "完成问号到箭头的听觉闭环，给片尾文字留呼吸", "M06", "笔尖连续划线、合本、教室底噪、尾部静默"),
]


shots = []
for i, item in enumerate(shot_content, start=1):
    start = (i - 1) * 10
    end = i * 10
    plot, visual, state, function, need, mtype, emotion, enter, exit_, volume, effect, cue, sfx = item
    shots.append({
        "镜头": f"{i:02d}",
        "时间": f"{tc(start)}-{tc(end)}",
        "剧情内容": plot,
        "视觉情绪": visual,
        "人物状态": state,
        "剧情作用": function,
        "是否需要音乐": need,
        "需要类型": mtype,
        "音乐情绪": emotion,
        "音乐进入时间": enter,
        "音乐结束时间": exit_,
        "音量": volume,
        "音乐作用": effect,
        "音乐名称": cue,
        "来源": "AI原创主方案" if cue.startswith("M") else ("片内声" if cue == "DIEGETIC" else "无配乐"),
        "音效重点": sfx,
    })


cues = [
    {
        "编号": "M01", "名称": "一个问号", "适用镜头": "04-07", "时长": "40秒（含2秒头尾把手）", "BPM": "58", "调性": "D小调/开放五度",
        "结构": "0-5秒空气与单音；5-24秒四音问号动机；24-34秒加入极轻大提琴泛音；34-40秒留尾并可与M02交叉",
        "Prompt": "电影级中国现实主义人文短片配乐，极简新古典，冷灰蓝课堂质感，58 BPM，D小调与开放五度，毡音钢琴以四音短动机表达疑问与克制，极轻空气合成器和大提琴泛音，稀疏、留白、无明确鼓点。情绪从疏离和疑问微微转向行动意愿，不悲情，不悬疑惊吓。40秒，无歌词，无人声，无可辨识现有旋律；前后各保留2秒剪辑把手；输出主混音及钢琴、弦乐、氛围分轨，48kHz/24bit WAV。禁止：宏大铜管、合唱、流行鼓、广告感、过度混响、突然高潮、华丽炫技。"
    },
    {
        "编号": "M02", "名称": "走出去", "适用镜头": "07-12", "时长": "56秒（含把手）", "BPM": "72", "调性": "D多利亚",
        "结构": "0-10秒从M01尾音接入；10-32秒木质轻脉冲支持步行；32-45秒逐步抽掉节奏进入老社区；45-56秒滤暗并把节拍交给挂钟",
        "Prompt": "写实纪实电影行进配乐，72 BPM，D多利亚调式，延续同一四音主旋律，毡音钢琴、极轻木质敲击、柔和拨弦和低频空气垫，表现女大学生带着问题从教室走向社区。节奏稳定但克制，具有探索和轻微期待，不要轻快Vlog感。末段逐步抽掉脉冲、频谱变暗，为木门和挂钟滴答让位。56秒，无歌词，无人声，前后2秒剪辑把手，48kHz/24bit WAV，输出无打击乐版本。禁止：企业宣传片和弦、尤克里里、拍手、强鼓、电子舞曲、明亮口哨、现成歌曲旋律。"
    },
    {
        "编号": "M03", "名称": "旧物会说话", "适用镜头": "13-19", "时长": "74秒（含把手）", "BPM": "56", "调性": "A小调，末段转C大调色彩",
        "结构": "0-15秒物件与空间；15-42秒大提琴泛音承托口述；42-56秒钥匙落桌后留一拍；56-68秒四音动机补全；68-74秒打开高频通向户外",
        "Prompt": "电影级人文口述史配乐，56 BPM，A小调但不悲情，温暖旧木屋、旧相册和黄铜钥匙的触感。毡音钢琴、独奏大提琴泛音、中低弦长音、极少玻璃马林巴点音与柔和模拟空气层；大段留白给老人对白、翻页、挂钟和钥匙实体声。后半段由小调缓慢出现C大调色彩，表达历史成为可触摸的生活证据。74秒，无歌词无人声，无民族符号化旋律，不模仿具体作曲家或作品；输出主混音、钢琴、弦乐、氛围分轨及无旋律对白版，48kHz/24bit WAV。禁止：苦情煽情、二胡主奏、战鼓、战争氛围、厚重低音轰鸣、合唱、突发转折、磁带故障噪声。"
    },
    {
        "编号": "M04", "名称": "院子里的答案", "适用镜头": "20-24", "时长": "54秒（含把手）", "BPM": "68", "调性": "G大调带Lydian色彩",
        "结构": "0-12秒庭院展开；12-30秒降低密度让对白和脚步主导；30-44秒四音主题完整展开；44-54秒回收为钢琴与大提琴并留白",
        "Prompt": "温暖但克制的现实主义电影配乐，68 BPM，G大调带轻微Lydian明亮色彩，柔和弦乐、小音量毡音钢琴、单簧管或中提琴轻奏四音主题，表现老社区无障碍坡道、电梯、阅读室和树下长椅所体现的生活改善。和声由中性缓慢打开，保持真实、踏实、有人情味，不像宣传片。54秒，无歌词无人声，动态中弱，给对白留频段和停顿；输出弦乐、钢琴、木管、氛围分轨，48kHz/24bit WAV。禁止：胜利号角、铜管齐奏、军鼓、史诗合唱、过度昂扬、流行鼓、企业励志音乐、快速蒙太奇感。"
    },
    {
        "编号": "M05", "名称": "钥匙与答案", "适用镜头": "25-30", "时长": "64秒（含把手）", "BPM": "64渐进至76体感", "调性": "D小调转F大调并开放收束",
        "结构": "0-10秒钥匙交接后进入；10-30秒主题由低八度转明亮；30-42秒对白处抽空；42-54秒为自拍视频蓄力；54-60秒路灯亮时达到全片唯一和声峰值；60-64秒迅速回落",
        "Prompt": "本片情感高潮配乐，电影级克制新古典，64 BPM并通过分解音型形成76 BPM的前进体感，D小调逐步转F大调。毡音钢琴重复四音问号主题，独奏大提琴先低八度回应，柔和室内弦乐逐层加入；在老人回答处主动抽空，在女主自拍视频最后一句与背景路灯点亮后才达到全片唯一峰值，随后立即回落。表达代际传承、理解、行动与希望，不是煽情颁奖音乐。64秒，无歌词无人声，无现成旋律；输出全混、钢琴、大提琴、弦乐、无打击乐、对白稀疏版，48kHz/24bit WAV。禁止：铜管、合唱、重鼓、电影预告片risers、轰鸣、过度悲伤、英雄主义、歌唱性过强旋律。"
    },
    {
        "编号": "M06", "名称": "带回来", "适用镜头": "31-36", "时长": "65秒（含把手）", "BPM": "68", "调性": "D大调add9",
        "结构": "0-18秒明亮回奏主题；18-20秒淡出；20-30秒预留片内投影声，不提供外部配乐；30秒后钢琴再入；30-52秒弦乐缓慢上扬；52-59秒箭头动作完成主题；59-65秒合本后自然尾音",
        "Prompt": "现实主义微电影结尾回归主题，68 BPM，D大调add9，重现开场四音问号动机但最后一音稳定落地，毡音钢琴、温暖低弦和克制高弦。0-18秒支持女主返回课堂；18-20秒淡出；20-30秒设计为近乎完全静默的十秒窗口，供投影片内对白和片内音乐播放；30秒后钢琴在教师转头之后再入；末段随“走出去、带回来”和问号画成箭头逐步上扬，合本后只留自然尾音并在最后1.5秒归于room tone。65秒，无歌词无人声，48kHz/24bit WAV，输出完整结构及30秒后独立可编辑尾段。禁止：毕业典礼感、颁奖音乐、铜管、合唱、强鼓、强行大团圆、尾部重击、过度延长混响。"
    },
]


resources = [
    {"名称": "Solace", "作者": "Scott Buckley", "平台": "Scott Buckley Creative Commons Music Library", "链接": "https://www.scottbuckley.com.au/library/solace/", "类型": "沉思钢琴、弦乐、氛围合成器、玻璃马林巴", "适用镜头": "04-13，作为M01/M02备选", "版权状态": "CC BY 4.0；官方曲目页已核验", "商业使用要求": "可商用及改编，必须按作者格式署名；YouTube署名放视频说明；不得将音乐单独转售/重发到流媒体，不得送入Content ID；无法署名需购买许可"},
    {"名称": "Hiraeth", "作者": "Scott Buckley", "平台": "Scott Buckley Creative Commons Music Library", "链接": "https://www.scottbuckley.com.au/library/hiraeth/", "类型": "安静合成器、弦乐、独奏大提琴，怀旧而克制", "适用镜头": "13-19，作为M03备选", "版权状态": "CC BY 4.0；官方曲目页已核验", "商业使用要求": "可商用及改编，必须署名；同步于影片使用；不得独立分发或做音频指纹；无署名用途需购买许可"},
    {"名称": "There Was A Time", "作者": "Scott Buckley", "平台": "Scott Buckley Creative Commons Music Library", "链接": "https://www.scottbuckley.com.au/library/there-was-a-time/", "类型": "弦乐与合成器，怀旧中带轻微上扬", "适用镜头": "14-24，作为M03/M04备选", "版权状态": "CC BY 4.0；官方曲目页已核验", "商业使用要求": "可商用及剪辑，必须署名；不得把音乐作为独立商品或上传音乐平台；不得提交Content ID"},
    {"名称": "A Kind Of Hope", "作者": "Scott Buckley", "平台": "Scott Buckley Creative Commons Music Library", "链接": "https://www.scottbuckley.com.au/library/a-kind-of-hope/", "类型": "苦甜钢琴、弦乐与氛围合成器", "适用镜头": "25-30或31-36，作为M05/M06备选", "版权状态": "CC BY 4.0；官方曲目页已核验", "商业使用要求": "可商用及改编，必须按官方格式署名；YouTube说明栏署名；不得独立分发/音频指纹；不便署名需购买项目许可"},
    {"名称": "Resolutions", "作者": "Scott Buckley", "平台": "Scott Buckley Creative Commons Music Library", "链接": "https://www.scottbuckley.com.au/library/resolutions/", "类型": "钢琴与弦乐，回望过去并面向未来", "适用镜头": "31-36，作为M06备选；只选低密度段落", "版权状态": "CC BY 4.0；官方曲目页已核验", "商业使用要求": "可商用及改编，必须署名；不得独立转售/重发，不得提交Content ID；无署名版本需购买许可"},
]


sfx = [
    {"音效名称": "教室空间底噪与窗外树叶", "对应镜头": "01-07、32-36", "来源": "剧组现场双系统录音；同机位采60秒room tone", "是否需要生成": "否，必须实录", "执行": "S01冷白下午与S05蓝调傍晚分别录；无对白处保持-34至-30 dBFS感知底"},
    {"音效名称": "粉笔书写与粉笔放下", "对应镜头": "02", "来源": "后期拟音，使用同材质黑板与粉笔", "是否需要生成": "否", "执行": "近拾音但削弱2-4kHz刺耳区；问号最后一笔清楚"},
    {"音效名称": "笔记本、笔尖、合书、包布与桌椅", "对应镜头": "03-07、18、27、35-36", "来源": "道具原物拟音", "是否需要生成": "否", "执行": "红色硬壳本建立固定声音身份；镜头36合本声不可被尾奏盖住"},
    {"音效名称": "走廊/校园/巷道三组脚步", "对应镜头": "08-10、19-21、29、31", "来源": "按鞋底和地面分别拟音", "是否需要生成": "否", "执行": "白帆布鞋在灰砖、旧巷、教室地面分别录；远近与机位匹配"},
    {"音效名称": "远处自行车铃", "对应镜头": "09", "来源": "现场实录或已购商业音效库单次铃声", "是否需要生成": "否", "执行": "只出现一次，置于远景右后方，避免喜剧感"},
    {"音效名称": "木门、门轴与门槛", "对应镜头": "11-12、19", "来源": "S03同一木门实录", "是否需要生成": "否", "执行": "同一门使用同一音色；进入与离开用透视和混响区分"},
    {"音效名称": "老挂钟滴答", "对应镜头": "12-18", "来源": "现场实录；必要时用合法商用库补录", "是否需要生成": "否", "执行": "作为室内时间感，不与音乐节拍完全量化；对白时降6 dB"},
    {"音效名称": "相册翻页、杯底、木桌摩擦", "对应镜头": "13-16", "来源": "近距离拟音", "是否需要生成": "否", "执行": "保持旧纸、布面相册和实木桌的干燥质感"},
    {"音效名称": "黄铜钥匙与红棉绳", "对应镜头": "17、25-28", "来源": "P06原物拟音", "是否需要生成": "否", "执行": "建立统一金属音色；交接以红绳摩擦为主，避免夸张清脆叮声"},
    {"音效名称": "稳定器电机、录制键与支架落地", "对应镜头": "07、16、22、29、32", "来源": "P03原机实录", "是否需要生成": "否", "执行": "设备声轻、短、真实；禁止科幻UI音"},
    {"音效名称": "庭院风、梧桐叶、鸟鸣与远处生活声", "对应镜头": "20-30", "来源": "现场多轨氛围录音", "是否需要生成": "可选，仅缺失时生成连续底层", "执行": "至少录90秒无缝底；不出现可辨人名或完整谈话；鸟鸣最多2次"},
    {"音效名称": "坡道脚步与手擦金属扶手", "对应镜头": "21", "来源": "同步实录+拟音补强", "是否需要生成": "否", "执行": "三步清楚但不过分放大，金属摩擦短促"},
    {"音效名称": "外装电梯低频运行底", "对应镜头": "22-24", "来源": "现场远距实录", "是否需要生成": "否", "执行": "若画面未显示运行则只保留极弱建筑机电底，不添加提示铃"},
    {"音效名称": "路灯继电器轻响", "对应镜头": "30", "来源": "拟音", "是否需要生成": "否", "执行": "极轻，仅作灯亮同步点，不做魔法音效"},
    {"音效名称": "投影连接提示、风扇与小扬声器播放质感", "对应镜头": "32-35", "来源": "现场设备实录与EQ处理", "是否需要生成": "否", "执行": "片内声做150Hz-8kHz带宽限制并加教室短混响；镜头33禁止叠加外部配乐"},
]


checks = [
    ("音乐风格统一", "通过", "六首原创Cue共享四音问号动机、毡音钢琴和克制弦乐；色彩随冷灰蓝→暖金→蓝金变化。"),
    ("人物主题一致", "通过", "林悦由未完成四音动机到结尾稳定落音；周建国不设独立英雄主题，以大提琴与物件声代表生活记忆。"),
    ("剧情变化匹配", "通过", "静默提问→行进→口述史→现实空间→钥匙交接高潮→课堂回归，动态只在镜头30形成一次峰值。"),
    ("关键情绪覆盖", "通过", "镜头04、17、22、25、26、28、30、34、36均有明确音乐或静默策略。"),
    ("无意义铺乐", "通过", "镜头01-03无配乐；镜头33不叠外部配乐；多处让对白、粉笔、翻页、钥匙、脚步主导。"),
    ("音乐重复控制", "通过", "主题重复以调性、配器、密度和功能变奏；不直接循环同一整轨。"),
    ("版权风险", "有条件通过", "主方案为项目原创AI音乐，需归档所用平台条款、生成记录与下载凭证；备选曲均来自作者官方页并标注CC BY 4.0。"),
    ("对白可懂度", "通过", "所有对白镜头规定duck 4-8 dB；镜头28和34以近静默处理。"),
    ("技术交付", "待制作时执行", "48kHz/24bit WAV、分轨、2秒把手、最终-16 LUFS±1、True Peak≤-1 dBTP。"),
]


data = {
    "项目": "《把答案写在大地上》",
    "版本": "音乐设计 v1.0",
    "日期": "2026-09-18",
    "成片": "6分00秒，36镜头，每镜头10秒，24fps",
    "默认执行": "AI原创六Cue连续组曲；版权库曲目只作备选，不与原创方案混用",
    "shots": shots,
    "cues": cues,
    "resources": resources,
    "sfx": sfx,
    "checks": checks,
}


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color="D9D9D9", size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        node = borders.find(qn(tag))
        if node is None:
            node = OxmlElement(tag)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def keep_with_next(paragraph):
    paragraph.paragraph_format.keep_with_next = True


def add_table(doc, headers, rows, widths=None, font_size=8.5):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    hdr = table.rows[0]
    set_repeat_table_header(hdr)
    prevent_row_split(hdr)
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = str(h)
        set_cell_shading(cell, "1F4E78")
        set_cell_border(cell)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = "Microsoft YaHei"
                run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                run.font.size = Pt(font_size)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
    for r_idx, row in enumerate(rows):
        added_row = table.add_row()
        prevent_row_split(added_row)
        cells = added_row.cells
        for i, value in enumerate(row):
            cells[i].text = str(value)
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_border(cells[i])
            if r_idx % 2 == 1:
                set_cell_shading(cells[i], "F3F6FA")
            for p in cells[i].paragraphs:
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.05
                if len(str(value)) < 14 and i != 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = "Microsoft YaHei"
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
                    run.font.size = Pt(font_size)
        if widths:
            for i, width in enumerate(widths):
                cells[i].width = Cm(width)
    if widths:
        for i, width in enumerate(widths):
            hdr.cells[i].width = Cm(width)
    return table


def add_label_para(doc, label, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_together = True
    r = p.add_run(label)
    r.bold = True
    r.font.name = "Microsoft YaHei"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    p.add_run(text)
    return p


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def style_document(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(25, 25, 25)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.2
    for name, size, bold in [("Title", 24, True), ("Heading 1", 16, True), ("Heading 2", 13, True), ("Heading 3", 11, True)]:
        st = styles[name]
        st.font.name = "Microsoft YaHei"
        st._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        st.font.size = Pt(size)
        st.font.bold = bold
        st.font.color.rgb = RGBColor(0, 0, 0)
        st.paragraph_format.space_before = Pt(10)
        st.paragraph_format.space_after = Pt(6)
        st.paragraph_format.keep_with_next = True


def build_docx():
    doc = Document()
    style_document(doc)
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.6)
    sec.left_margin = Cm(1.9)
    sec.right_margin = Cm(1.9)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.8)
    footer = sec.footer.paragraphs[0]
    footer.text = "《把答案写在大地上》音乐设计方案  ·  "
    add_page_number(footer)

    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("把答案写在大地上\n整体音乐设计方案")
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run("镜头级配乐 音效与版权执行版\n版本 v1.0  2026年9月18日").italic = True
    doc.add_paragraph()
    opening = doc.add_paragraph()
    opening.add_run("执行结论：").bold = True
    opening.add_run("本片采用六首共享同一四音动机的AI原创配乐组成连续组曲。前33.5秒保持无配乐，高潮只出现在镜头30，镜头33禁止叠加画外配乐。音乐从“问号”出发，经旧相册与钥匙完成生活证据的触觉化，最终在问号变为箭头时稳定落地。版权库曲目仅作为无法完成原创制作时的备选，不与原创方案混用。")
    add_table(doc, ["项目参数", "执行值"], [
        ("片名与用途", "《把答案写在大地上》；高校思政课展映、校园微电影比赛、学校官方视频平台"),
        ("画面结构", "6分00秒；36镜头；每镜头10秒；写实纪实；冷灰蓝→暖金→蓝金"),
        ("声音核心", "对白与生活实声优先；音乐服务“疑问—观察—触动—理解—行动”"),
        ("交付规格", "音乐48kHz/24bit WAV；主混+分轨；前后各2秒把手；最终建议-16 LUFS±1，True Peak≤-1 dBTP"),
    ], widths=[4.0, 12.5], font_size=9)

    doc.add_heading("第一部分 整体音乐设计方案", level=1)
    doc.add_heading("整体音乐风格", level=2)
    doc.add_paragraph("克制的新古典电影配乐与纪实环境声融合。核心音色是毡音钢琴、独奏大提琴、柔和室内弦乐、极轻木质脉冲和低密度空气合成器。音乐不做宏大宣传片式陈述，也不把旧社区段处理成苦情回忆；它只在人物的观看方式发生变化时增加信息。")
    add_label_para(doc, "统一听觉母题：", "四音“问号动机”，前段停在不稳定音，结尾才稳定落地。每次重现改变调性、配器或密度，形成统一而不重复的听觉叙事。")
    add_label_para(doc, "主旋律 Theme：", "M01提出四音问号；M02变成步行脉冲；M03在旧相册段以大提琴和钢琴碎片化呈现；M04转为明亮和声；M05在钥匙交接后完成；M06以大调add9收束。")
    add_label_para(doc, "人物主题 Character Theme：", "林悦的主题就是问号动机，其从未完成到落地的过程代表她由标准答案思维转向社会观察。陈老师不用独立主题，以清晰静默和钢琴留白代表方法引导。周建国不用英雄主题，以大提琴、旧木桌、相册与钥匙实体声构成生活记忆主题。")
    add_label_para(doc, "冲突主题 Conflict Theme：", "本片没有外部对抗，冲突来自“书本定义”与“生活证据”的距离。用开放五度、未解决音与音乐突然抽空表现，而不是用紧张鼓点。")
    add_label_para(doc, "高潮主题 Climax Theme：", "镜头25钥匙交接启动，镜头26完成内在转折，镜头30最后一句与路灯亮起达到唯一和声峰值。镜头28的核心对白反而近静默，保证思想内容清楚。")

    doc.add_heading("整体配乐方向", level=2)
    add_table(doc, ["维度", "设计"], [
        ("乐器", "毡音钢琴、独奏大提琴、室内弦乐、中提琴/单簧管、木质轻敲、低密度空气合成器。禁止铜管齐奏、史诗合唱、军鼓、尤克里里和广告式拍手。"),
        ("速度", "56-72 BPM为主；M05通过分解音型形成76 BPM体感，真正节拍仍克制。"),
        ("节奏", "前段无拍或单音；出发段有轻脉冲；口述段让挂钟和物件声接管时间；庭院段拉长和声；高潮以层次增加而非鼓点增速。"),
        ("情绪变化", "疏离疑问→行动探索→怀旧触摸→温暖看见→代际传承→笃定回归。"),
        ("音色颜色", "冷灰蓝用偏干钢琴与空气层；木屋用木质中低频与大提琴；暖金庭院用中弦和柔木管；蓝金教室用钢琴回归与温暖低弦。"),
        ("对白原则", "对白镜头音乐自动duck 4-8 dB；重点句前后留0.3-1秒呼吸。镜头33只播片内声，禁止双层配乐。"),
    ], widths=[3.2, 13.3], font_size=9)

    doc.add_heading("连续剧情音乐结构", level=2)
    add_table(doc, ["比例/时间", "剧情阶段", "音乐结构", "主要Cue"], [
        ("0%-20%\n00:00-01:12", "课堂提出问题并出发", "前33.5秒无配乐；问号动机单音进入；末段转为步行脉冲", "SIL / M01 / M02"),
        ("20%-50%\n01:12-03:00", "进入社区与口述历史", "脉冲逐渐抽掉；挂钟接管时间；相册与钥匙触发低声弦乐", "M02 / M03"),
        ("50%-80%\n03:00-04:48", "现实设施与认知转折", "声场打开；和声由中性转暖；镜头28对白处近静默", "M03 / M04 / M05"),
        ("80%-100%\n04:48-06:00", "自拍视频高潮与回到课堂", "镜头30唯一峰值；回归主题；镜头33外部音乐停；问号变箭头后自然消散", "M05 / M06 / 片内声"),
    ], widths=[3.0, 4.0, 6.7, 2.8], font_size=8.7)

    doc.add_heading("Cue总览", level=2)
    add_table(doc, ["编号", "名称", "镜头", "BPM", "调性", "功能"], [
        (c["编号"], c["名称"], c["适用镜头"], c["BPM"], c["调性"], c["结构"]) for c in cues
    ], widths=[1.3, 2.5, 2.0, 1.4, 3.0, 7.0], font_size=8.3)

    doc.add_heading("第二部分 镜头音乐需求分析", level=1)
    doc.add_paragraph("以下分析以最终确认的36张正式关键帧、完整剧本时间码与对白为准。音量为音乐总线建议值，实际混音需以对白清晰度和最终响度目标复核。")
    doc.add_page_break()
    for s in shots:
        shot_no = int(s["镜头"])
        if shot_no > 1 and (shot_no - 1) % 3 == 0:
            doc.add_page_break()
        doc.add_heading(f"镜头{s['镜头']}  {s['时间']}", level=2)
        rows = [
            ("剧情内容", s["剧情内容"]),
            ("视觉情绪 / 人物状态", f"{s['视觉情绪']}；{s['人物状态']}"),
            ("剧情作用", s["剧情作用"]),
            ("音乐需求", f"{s['是否需要音乐']}；{s['需要类型']}"),
            ("音乐情绪", s["音乐情绪"]),
            ("进入 / 结束", f"{s['音乐进入时间']} / {s['音乐结束时间']}"),
            ("音量", s["音量"]),
            ("音乐作用", s["音乐作用"]),
            ("音效重点", s["音效重点"]),
        ]
        add_table(doc, ["项目", "镜头执行说明"], rows, widths=[3.3, 13.2], font_size=8.8)

    doc.add_heading("第三部分 AI音乐生成Prompt", level=1)
    doc.add_paragraph("六首Cue必须使用同一生成平台、同一音色参考与同一四音主题种子。若平台不支持分轨，应至少生成主混音、无打击乐版、对白稀疏版；生成后由人工剪辑和混音，不直接把未经审听的AI结果上片。")
    for c in cues:
        doc.add_heading(f"{c['编号']}  {c['名称']}", level=2)
        add_label_para(doc, "适用镜头：", c["适用镜头"])
        add_label_para(doc, "速度与调性：", f"{c['BPM']} BPM；{c['调性']}")
        add_label_para(doc, "结构：", c["结构"])
        add_label_para(doc, "生成Prompt：", c["Prompt"])

    doc.add_heading("第四部分 版权音乐资源推荐", level=1)
    doc.add_paragraph("以下曲目仅作为原创制作失败、赛前临时替换或导演试剪的备选。曲目页与许可证要求于2026年9月18日从作者官方音乐库核验。正式采用前仍需保存曲目页截图、下载日期、下载文件和署名文本。")
    for idx, r in enumerate(resources):
        doc.add_heading(f"{r['名称']}  {r['作者']}", level=2)
        rows = [("推荐平台", r["平台"]), ("官方链接", r["链接"]), ("音乐类型", r["类型"]), ("适用镜头", r["适用镜头"]), ("版权状态", r["版权状态"]), ("商业使用要求", r["商业使用要求"])]
        add_table(doc, ["项目", "信息"], rows, widths=[3.3, 13.2], font_size=8.8)
    add_label_para(doc, "统一署名格式：", "“‘Track Title’ by Scott Buckley – released under CC-BY 4.0. www.scottbuckley.com.au”。若上传YouTube，署名必须放在视频说明栏；片尾字幕可同步保留。")
    add_label_para(doc, "官方许可说明：", "https://www.scottbuckley.com.au/library/using-this-music/；CC BY 4.0：https://creativecommons.org/licenses/by/4.0/")

    landscape = doc.add_section(WD_SECTION.NEW_PAGE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = sec.page_height, sec.page_width
    landscape.top_margin = Cm(1.4)
    landscape.bottom_margin = Cm(1.4)
    landscape.left_margin = Cm(1.2)
    landscape.right_margin = Cm(1.2)
    footer2 = landscape.footer.paragraphs[0]
    footer2.text = "《把答案写在大地上》镜头音乐对应表  ·  "
    add_page_number(footer2)
    doc.add_heading("第五部分 镜头音乐对应表", level=1)
    mapping_rows = []
    for s in shots:
        mapping_rows.append((s["镜头"], s["时间"], s["剧情内容"], s["音乐名称"], s["来源"], s["音乐进入时间"], s["音乐结束时间"], s["音量"], s["音乐作用"]))
    add_table(doc, ["镜头", "时间", "画面", "音乐名称", "来源", "进入点", "退出点", "音量", "作用"], mapping_rows,
              widths=[1.0, 1.8, 5.0, 2.0, 2.2, 1.8, 1.8, 3.0, 6.0], font_size=7.1)

    portrait = doc.add_section(WD_SECTION.NEW_PAGE)
    portrait.orientation = WD_ORIENT.PORTRAIT
    portrait.page_width, portrait.page_height = sec.page_width, sec.page_height
    portrait.top_margin = Cm(1.8)
    portrait.bottom_margin = Cm(1.6)
    portrait.left_margin = Cm(1.9)
    portrait.right_margin = Cm(1.9)
    footer3 = portrait.footer.paragraphs[0]
    footer3.text = "《把答案写在大地上》音乐设计方案  ·  "
    add_page_number(footer3)
    doc.add_heading("第六部分 音效设计方案", level=1)
    doc.add_paragraph("声音设计坚持“生活实声先于特殊音效”。粉笔、脚步、门、相册、钥匙、稳定器、投影设备构成可辨识的物件声链；不使用战争、幻想、科幻和宣传片式whoosh。")
    add_table(doc, ["音效名称", "对应镜头", "来源", "是否生成", "执行说明"], [(x["音效名称"], x["对应镜头"], x["来源"], x["是否需要生成"], x["执行"]) for x in sfx], widths=[3.4, 2.6, 3.6, 2.0, 5.0], font_size=7.8)
    doc.add_heading("声音空间与混音规则", level=2)
    for text in [
        "对白为第一优先级。对白出现时音乐自动duck 4-8 dB；S03室内对白使用短小木屋反射，S04保持开阔空气感，S05片内投影声做150Hz-8kHz带宽限制。",
        "环境声跨镜连续。走廊、巷道、木屋、庭院、教室各保留至少60-90秒无缝room tone；硬切画面时可让环境声提前或延后4-12帧形成J-cut/L-cut。",
        "钥匙的金属音必须统一。不要每次使用不同库音；优先用P06原物录制一组拿起、落桌、手心交接、放在本子和红绳摩擦。",
        "最终混音建议：对白中心稳定；音乐不占2-4kHz；成片综合响度-16 LUFS±1，True Peak不高于-1 dBTP。课堂放映版另做动态略收窄的备份，不改变母版。",
    ]:
        doc.add_paragraph(text, style=None)

    doc.add_heading("第七部分 音乐制作检查报告", level=1)
    add_table(doc, ["检查项", "结论", "说明"], checks, widths=[3.4, 2.3, 10.0], font_size=8.5)
    doc.add_heading("制作前必须归档", level=2)
    for item in [
        "AI平台名称、账号主体、生成日期、Prompt、模型版本、导出文件哈希与当日商业使用条款截图。",
        "每首版权库音乐的官方曲目页、许可证页、下载日期、原始文件、署名文本和任何购买发票。",
        "音乐Cue表、版本号、剪辑使用段、变速/变调/剪切记录、最终片尾和平台说明栏署名截图。",
        "所有现场录音与拟音的录制人、日期、设备、原始文件和工程项目，确保权属可追溯。",
    ]:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(item)
    doc.add_heading("最终结论", level=2)
    doc.add_paragraph("方案已覆盖36个镜头、六个连续剧情Cue、15组音效资产和5首可追溯版权备选曲。首选执行为AI原创六Cue组曲；若改用版权曲，必须整段替换对应Cue并完成署名，不能把五首备选随意拼贴。后期可直接以配套Excel表筛选镜头、Cue、进出点、音量与音效重点。")

    doc.save(DOCX_PATH)


def write_supporting_files():
    JSON_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    README_PATH.write_text(
        "# 《把答案写在大地上》音乐设计交付\n\n"
        "本文件夹是6分钟、36镜头微电影的配乐与声音执行包。\n\n"
        "## 文件\n\n"
        "- `《把答案写在大地上》整体音乐设计方案.docx`：完整七部分方案，适合导演、配乐师和声音后期审阅。\n"
        "- `《把答案写在大地上》镜头-音乐对应关系表.xlsx`：后期直接筛选和执行的工作表，包含镜头映射、Cue、AI Prompt、版权资源、音效和检查项。\n"
        "- `《把答案写在大地上》镜头音乐数据.json`：结构化数据，便于导入剪辑/制片工具。\n"
        "- `版权署名模板.txt`：采用CC BY 4.0备选曲时直接复制并替换曲名。\n\n"
        "## 默认执行决定\n\n"
        "优先制作M01-M06六首AI原创Cue。前33.5秒无配乐，镜头30为唯一和声峰值，镜头33不叠加外部配乐。版权曲只作为整段替代方案，不与原创Cue混搭。\n\n"
        "## 技术基线\n\n"
        "音乐与音效统一48kHz/24bit WAV；Cue保留前后2秒把手；建议成片-16 LUFS±1、True Peak不高于-1 dBTP。\n",
        encoding="utf-8",
    )
    ATTR_PATH.write_text(
        "《把答案写在大地上》版权音乐署名模板\n\n"
        "片尾/项目说明：\n"
        "Music: '[Track Title]' by Scott Buckley – released under CC-BY 4.0. www.scottbuckley.com.au\n\n"
        "YouTube说明栏：\n"
        "'[Track Title]' by Scott Buckley – released under CC-BY 4.0. www.scottbuckley.com.au\n"
        "Track page: [粘贴官方曲目链接]\n"
        "License: https://creativecommons.org/licenses/by/4.0/\n\n"
        "执行提示：不得将曲目单独转售、重发到音乐流媒体或提交Content ID。不能公开署名时，需向作者购买对应项目许可。\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    write_supporting_files()
    build_docx()
    print(DOCX_PATH)
