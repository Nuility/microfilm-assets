# 统一视觉规范与Negative Prompt

## 统一正向风格尾缀

当代中国现实主义剧情片，电影级写实摄影，真实影视剧质感，自然皮肤纹理和真实织物材质，生活化自然表演，真实空间比例，克制浅景深，电影构图，细节清晰但不过度锐化，无磨皮塑料感，16:9横屏，无文字、无品牌、无水印。

## 统一Negative Prompt

different face, different character, identity drift, wrong age, age change, different hairstyle, hairstyle change, different clothes, clothes color change, wrong body proportion, accessories missing, extra accessories, skin tone change, extra people, crowd, extra objects, prop duplication, prop disappearance, wrong anatomy, bad hands, extra fingers, missing fingers, fused fingers, deformed body, unrealistic face, asymmetrical eyes, distorted glasses, floating objects, impossible pose, incorrect perspective, warped architecture, changed room layout, changed building facade, inconsistent light direction, sudden weather change, rain, overexposure, crushed shadows, cartoon, anime, illustration, painting style, plastic skin, beauty filter, excessive HDR, low quality, blurry, compression artifacts, inconsistent style, readable text, garbled text, subtitle, logo, school emblem, brand, watermark, political slogan.

## A01专项Negative Prompt

generic face, ordinary bland face, influencer face, internet celebrity makeup, doll eyes, oversized eyes, heavy double eyelids, heavy makeup, false eyelashes, glossy lips, porcelain skin, excessive skin whitening, mature glamorous woman, fragile weak posture, hunched shoulders, fashion advertisement pose, high ponytail, straight bangs, curled hair, missing beauty mark, beauty mark on wrong side, long coat, oversized coat, skirt, high heels.

## 手部专项Negative Prompt

extra hand, third hand, duplicated hand, extra fingers, missing fingers, fused fingers, twisted wrist, broken finger, object passing through hand, floating object, incorrect left-right hand, malformed fingernails.

## 双人专项Negative Prompt

third person, swapped identity, swapped clothing, incorrect height ratio, incorrect eyeline, merged bodies, overlapping limbs, changed standing position, duplicate character.

## 屏幕与文字专项Negative Prompt

readable UI, random Chinese characters, garbled text, brand interface, phone logo, subtitle, watermark, generated poster text, generated blackboard sentence, generated projection caption.

## 场景专项锁定

- S01/S05：白墙、浅木门、灰色吸音顶板、三列浅木课桌、中央讲台、右侧双窗、后方单门；不可改变布局。
- S02：米白墙、灰色地砖、金属栏杆；室外浅灰方砖步道与初秋墨绿树木；不可出现可识别人群。
- S03：1980年代红砖楼、局部修缮、深色木门；小屋内方木桌居中、左侧照片墙、右侧旧木窗、后方木柜。
- S04：中央梧桐树、左侧坡道和银色扶手、右侧阅读室、红砖楼外装电梯、树下长椅、固定单盏暖色路灯。
- 所有社区建筑保持同一红砖纹理、窗户比例和修缮程度。

## 使用规则

1. 每张生成图的完整Prompt后追加“统一正向风格尾缀”。
2. 每张图统一追加“统一Negative Prompt”。
3. 所有出现A01的画面追加“A01专项”；手部特写追加“手部专项”；双人画面追加“双人专项”；屏幕或本子画面追加“屏幕与文字专项”。
4. 参考图优先级：角色参考图 > 场景锚点图 > 前一镜连续性图。
5. 修改失败图时只改一个问题，其他内容明确写为保持不变。
