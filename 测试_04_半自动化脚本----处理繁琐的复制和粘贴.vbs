 
rem 
Dim wsh,a,li


li = Array("和平年代 + 蘑菇团","烟火 + HAMA陈缇","公路 - (Highway) + 纵贯线","真心英雄 (Live) + 纵贯线","灰色头像 + 许嵩","惊鸿一面 + 许嵩/黄龄","你若成风 + 许嵩/莫诗旎","幻听 + 许嵩","云淡风轻 + 王欣宇","胆小鬼 + 颜人中","超人 + 王贰浪","大悲咒 (梵唱) + 新韵传音","遇到 - (电视剧《恶作剧之吻》插曲) + 方雅贤","大悲咒 + 滇津喇嘛","最长的电影 + 茜拉","《起风了》三亚市2019届高三毕业生高考加油歌！（翻自 三亚学子）  + 唐有依/陈家伟/周晓巧/苏宽/许禄富/陈日威/林萌/兰恩飞/Shunger周爽/伍志伟/黄千禧","琵琶行 + 铃九Rin/沐闇","你曾是少年 + 焦迈奇","你就不要想起我 (live) - (原唱：田馥甄) + 张杰","仰望星空 + 张杰","心如止水 + Ice Paper","窗 - (电视剧《扶摇》片尾曲/人物主题曲) + 吴青峰","四块五 + 隔壁老樊","我的名字 + 焦迈奇","天后（Cover 薛之谦） + 张pd","让我留在你身边 - (原唱：陈奕迅) + 任嘉伦","星球坠落 (Live) + 艾热 AIR/李佳隆","目不转睛 + 王以太","篝火旁 + 吕大叶/马子林Broma/陈觅Lynne","2018小恋曲 + 沈虫虫/丫蛋蛋","说爱你 + 刘至佳","舞步 + 蔡健雅","天份 + 薛之谦","阳光彩虹小白马 + 大张伟","月牙湾 + 丫蛋蛋/沈虫虫","写给黄淮（demo） + 邵帅","不必说抱歉 + 胡66","夏风（男声） + 李欣龙啊","如果寂寞了（风情版） + 郑晓填","千禧 + 徐秉龙","幸福了 然后呢 + 黄丽玲","一个人去巴黎 + 董又霖","撒野 - (巫哲小说《撒野》官方主题曲) + 凯瑟喵","肆无忌惮 + 薛之谦","昨日青空 - (电影《昨日青空》同名青春主题曲) + 尤长靖","山海 + 隔壁老樊","你还要我怎样 + 薛之谦","遥不可及的你 + 花粥","远走高飞 + 金志文","浮生 + 刘莱斯","可乐 + 赵紫骅","大鱼 - (动画电影《大鱼海棠》印象曲) + 周深","月牙湾 + F.I.R.","罗生门 + 麦浚龙/谢安琪","慢慢喜欢你 - (Growing Fond of You) + 莫文蔚","特务J + 蔡依林","绝代风华 - (天下3十周年主题曲) + 许嵩","岁月神偷 + 金玟岐","拿走了什么 + 黄丽玲","天亮以前说再见 + 何野","你 + 田馥甄","我多喜欢你，你会知道 - (网剧《致我们单纯的小美好》推广曲) + 王俊琪","OK歌（翻自 麦小兜/单色凌） + 西二","偏爱 - (电视剧《仙剑奇侠传三》插曲) + 张芸京","消愁 + 毛不易","云烟成雨 - (动画《我是江小白》片尾曲) + 房东的猫","那些你很冒险的梦 - (Those Were The Days) + 林俊杰","给妈妈 - (电影《悲伤逆流成河》插曲) + 房东的猫","我好像在哪见过你 - (动画电影《精灵王座》主题曲) + 薛之谦","拜拜 + 浙音4811（一个大金意）","再见青春 - (电影《悲伤逆流成河》插曲) + 任素汐","水星记 + 郭顶","年少有为 + 李荣浩","毒苹果 + G.E.M.邓紫棋","你一定要幸福 (cover 何洁) + 简弘亦","你不是真正的快乐 (Live) + G.E.M.邓紫棋","默 - (电影《何以笙箫默》主题插曲) + 那英","怎么唱情歌Jazz + 柳戈","中毒 + 野区歌神","爱的魔法 + 阮豆","盛夏的果实 (Live) + 汪睿","我要你 - (电影《驴得水》主题曲) + 任素汐","世界上的另一个我 + 阿肆/郭采洁","广东十年爱情故事 + 广东雨神","东西 + 林俊呈","你就不要想起我 + 田馥甄","那一夜 + G.E.M.邓紫棋","其实，我就在你方圆几里（Cover 薛之谦） + 陈壹千","疑心病 + 任然","突然想起你(Live) + 林宥嘉","可惜没如果 - (If Only…) + 林俊杰","一生所爱 - (原唱：卢冠廷) + 莫文蔚","流着泪说分手 + 金志文","红色高跟鞋 - (电影《爱情呼叫转移2》主题曲) + 蔡健雅","时钟（特别版） + 丁芙妮","忽然之间 + 田新宇","我们俩（Cover 郭顶） + 叶齐心","卡路里 - (电影《西虹市首富》插曲) + 火箭少女101","侧脸 + 于果","那些你很冒险的梦 + Troly","坠落 + 蔡健雅","外愈 + 任然","鱼仔 - (电视剧《花甲男孩转大人》主题曲) + 卢广仲","从无到有 - (电视剧《创业时代》片尾曲) + 毛不易","寂寞烟火 + 朱婧汐Akini Jing","天地无霜 (合唱版) - (电视剧《香蜜沉沉烬如霜》对唱主题曲) + 杨紫/邓伦","左手指月 - (电视剧《香蜜沉沉烬如霜》片尾曲) + 萨顶顶","不要忘记我爱你 - (电视剧《神犬小七》主题曲) + 张碧晨","白羊 + 徐秉龙/沈以诚","不得不爱 + 潘玮柏/弦子","独行侠 + YKEY/Namunong","我在那一角落患过伤风 - (步步高音乐手机广告曲) + 冯曦妤","说谎 - (电影《针尖上的天使》主题曲) + 林宥嘉","气球 + 许哲珮","童话镇 + 暗杠","小半 + 陈粒","不染 - (电视剧《香蜜沉沉烬如霜》主题曲) + 毛不易","感谢你曾来过 + 周思涵/Ayo97","体面 + 千秋雪Candaramsi","再也没有 + 永彬Ryan.B/AY楊佬叁","红色高跟鞋 + 杜航","小奶狗VS大狼狗 + 郑建鹏/Y U Jay","最后我们没在一起 + 白小白","可不可以 + 张紫豪","淋雨一直走 - (Keep Walking) + 张韶涵","越来越不懂 + 蔡健雅","一百万个可能 - (A Million Possibilities) + Christine Welch","往后余生 + 王贰浪","再见 + G.E.M.邓紫棋","再见 你好 + 金玟岐","喜欢 + 阿肆","Shape of You×中文歌（翻自 Ed Sheeran） + 韩旭","给陌生的你听 + G.G张思源","后来的我们（翻自 五月天） + 二嘉","只想对你说 + 掌嘴/刘明宇Lil-7/张雪飞/IDEA MUSIC","隔壁泰山 + 阿里郎","烟火里的尘埃 + 华晨宇","你要的全拿走 - (Take Everything You Want) + 胡彦斌","浪人琵琶 + 胡66","青梅竹马 + 陈秋含","卡布奇诺 + 6诗人","睫毛弯弯 + 王心凌","最美的期待 - (电视剧《茧镇奇缘》主题曲) + 周笔畅","平凡之路 - (电影《后会无期》主题曲) + 朴树","有何不可（Cover 许嵩） + 锦零","痒 + 黄龄","BINGBIAN病变 (女声版) + 鞠文娴","我的一个道姑朋友（Cover タイナカ彩智 / Lon） + 纱琉璃Shelley","离人愁 + 李袁杰","纸短情长 (完整版) + 烟把儿","醉赤壁 - (网游《赤壁Online》主题曲) + 林俊杰","小酒窝 + 林俊杰/蔡卓妍","不仅仅是喜欢 + 萧全/孙语赛","后来 - (电视剧《真情》片尾曲) + 刘若英","明天，你好 - (电视剧《加油吧实习生》插曲) + 牛奶咖啡","春风吹 + 锦零","TheStar + 音阙诗听/李佳思","爱的就是你 + 刘佳","佛系少女 + 冯提莫","123我爱你 + 新乐尘符","红昭愿 + 音阙诗听","追光者 - (电视剧《夏至未至》插曲) + 岑宁儿","小公主 + 蒋蒋","非酋 + 薛黛霏/朱贺","情话微甜 + 王圣锋/李朝","东京不太热 + 封茗囧菌","电影情圣：我在人民广场吃炸鸡 + 程思佳","海草舞 + 萧全","小跳蛙 + 青蛙乐队","带你去旅行 + 校长","全部都是你 + DP龙猪/宝楽CNBALLER/CLOUDWANG 王云","演员 + 薛之谦","小幸运 - (电影《我的少女时代》主题曲) + 田馥甄","刚好遇见你 + 李玉刚","坐在巷口的那对男女 + 自然卷","豆花之歌（翻自 Pianoboy高至豪） + 锦零","夜空中最亮的星 + 逃跑计划","小情歌 + 苏打绿","说散就散 - (电影《前任3：再见前任》主题曲) + 袁娅维TIA RAY","空空如也 + 任然","9420 + 麦小兜","有点甜 - (《微微一笑很倾城》插曲) + 汪苏泷/By2","狐狸 - (电影《二代妖精之今生有幸》推广曲) + 薛之谦","远走高飞 (合唱版) + 金志文/徐佳莹","喜欢你 + G.E.M.邓紫棋","丑八怪 + 薛之谦","我们不一样 + 大壮","我的歌声里 + 曲婉婷","江南 - (River South) + 林俊杰","足够 + 周星星","白羊 + 苏源SVE","让我留在你身边 (原唱作版) - (Let Me Stay) + 唐汉霄","Daisy Blue + 鹿乃","明年今日 + 陈奕迅","BIGBANG + BIGBANG","遥远的她 + 陈奕迅","你的背包 + 陈奕迅","十年 + 陈奕迅","End of Time + Jim Yosef/Brenton Mattheus","Home + Vicetone","Rivers + Thomas Jack","飞云之下 + 韩红/林俊杰","不万能的喜剧 + 万能青年旅店","习惯 - (网剧《爱上北斗星男友》陪伴主题曲) + 周锐","帅到分手 - (电视剧《飞鱼高校生》片头曲) + 周汤豪","绕圈 + LeeAlive/盛哲","白羊 - (原唱：徐秉龙、沈以诚) + 曲肖冰","9277 + 贺知书/尚文婷","几度梦回大唐 + 深七","最美的太阳 + 张杰","这，就是爱 + 张杰","逆战 + 张杰","竹 (Bamboo) + Far East Movement/张杰","出山 + 花粥/王胜娚","我的秘密 - (My Secret) + G.E.M.邓紫棋","倒数 + G.E.M.邓紫棋","光年之外 - (电影《太空旅客》中文主题曲) + G.E.M.邓紫棋","一番の宝物 (Original Version) - (最为珍贵的宝物) + karuta","美しきもの - (美丽之物) + Sound Horizon","如烟 + 清风至","Broken + lovelytheband","小星星 + 汪苏泷","沙漠骆驼 + 展展与罗罗","Say Goodbye + 吴佳煜","打电话给你 + 秋秋","画一个星星一个你 + 小星星Aurora","宠坏 + 李俊佑/小潘潘（潘柚彤）","小鹿乱撞 + 永彬Ryan.B/狄迪（D-DAY）","1234喜欢你 + 王欣宇","习惯你的好 + 王理文","单人旁 + 汪昱名","照相机 + 王煜开 YK/舒涵Shellen","我愿意平凡的陪在你身旁 + 王七七","say you love me + 红宇乐团","有你就好 + 张津涤/倪可","生僻字 + 陈柯宇","迷途的羔羊 + 安子与九妹","塑料姐妹 + 盖巧","这么晚还没睡 + 孟想","无骨无花，无我无他 + 尚东峰","有可能的夜晚 + 曾轶可","灵魂伴侣 + 姜禹辰/w小婉君","空山新雨后 + 音阙诗听/锦零","讲真的 + 曾惜","不让你孤单 + 卓舒晨","猛吃猛喝 + 超暖/小猛哥","我可不可以 + DP龙猪","海市蜃楼 + 李瑨瑶","有幸 + 老光","确认过眼神 + 孙羽幽","夏焰 + 宋尚聪","你是我的 + 张欣尧","星空 + 纳豆nado","独特 + 过又嘉","作业狂想曲 + And2girls安菟","C A T (demo) + 曾婕Joey.Z","环游星空 + Gifty","古城一侧 + 毛儿/纸无笔砚","前方 + 第一乐章","元气满满 + 冯提莫","她的叮嘱 - (Her Words) + 朱家瑞","浪漫画面 + 冉小冉","错误直觉 + 刘贞岑","下坠Falling + Corki刘宗鑫","心跳 + 李予诺","大神带带我 + 菠萝赛东/新乐尘符","就是胖了怎么着 - (电影《月半爱丽丝》推广曲) + 陈柯宇","念念不如忘 + 樊桐舟","简简单单 + 丘沛宸","闺蜜 + 许嵩","关于分手 + Sevenking王七","心动 + DOUBLE","迷失飞船 (lost  ship) + 杨颠疯了","不曾相遇 + 亦勋","Warmer Winter.暖冬 + 木秦","Chu Desu! (feat. Dodogo) + Chunnyt","穿过整座城市 + Lil Chaos/Hayrul海力","和你很像的人 + 开开（赵锴羿）","怀里的猫 + 任舒瞳","小仙女 + 陈诗怡/萧全","花，我全部都给你 + 胡芳芳","叹郁孤 + 霄磊","合适的 + 张力超","闹/Now(Prod.by依兴驰) + N1FT/$UN","桃花妆 + 鸾音社","我们的征途是星辰和大海 + 王奇","因为你爱我 + 李鑫玲","GA GA + ANU","萌不萌 + 傅小爽","又是三更,念一个人 + 阿来","相爱慢慢 + 花与叔叔","意大利海 + 爱星人","彩虹 + 欧几/-艾兜","两只小猪 piggy piggy + 圈圈宝贝/蜜蜂少女队—DUDU&NINI","如前 + 张一然","青春小马 + 贰茉EMO","假装我们不需要呼吸 + 方拾贰（方十二）","窗外的颜色 + Gifty","只是曾经 + SunnyHouse弘哲","下个路口等我 + 吴思怡","元宵(feat.杭州童声) + 圈圈宝贝/圈圈小明星","倚风听雨 + 鸾音社","睡不着胸口闷 + 大川Dietry","男朋友女朋友 + 王嘉懿/阮豆","无条件爱你 + 爱星人","Money Girl + Taro芋头","解药 - (综艺《偶像练习生》首唱新歌) + BC221","晚安 + 留声玩具","着迷 + MAREA/wweimm/Xi3R","小圆 + 暗杠","置身事外 + 潘悦晨","可不可以给我你的微信 + 月光戦士","错就错了 + 姚慧/花草乐园会","直觉 + 张信哲","天空 + 王菲","月 + 好妹妹","光 + 陈粒","她 + 王威","小桥 + 暗杠","画 + 赵雷","路过人间 - (电视剧《我们与恶的距离》插曲) + 郁可唯","瞬 + 谢春花","蜂鸟 - (电视剧《我在北京等你》主题曲) + 吴青峰","学猫叫 + 小潘潘（潘柚彤）/小峰峰（陈峰）","你曾是少年 - (电影《少年班》主题曲) + S.H.E","狂浪 + 花姐","Way Back Home + SHAUN","生僻字(女生版) + 妧歌","起风了 - (电视剧《加油，你是最棒的》主题曲) + 吴青峰","i'm so tired... + Lauv/Troye Sivan","追光者 + 汪苏泷","欧若拉 (Live) + 张韶涵","我的滑板鞋2016 - (原唱：约瑟翰庞麦郎) + 华晨宇","不息之河 + TFBOYS","Stay Here Forever（Cover Jewel） + Britneylee小暖","Bad Girl（翻自 吴亦凡） + 卢润泽","Day by Day (Ver. Heartbreak) + BIGBANG","BLUE (Live) + BIGBANG","Baby Baby - (《最后的问候》英文版) + BIGBANG","HARU HARU - (一天一天) + BIGBANG","Baby(Acoustic Version) + Tyler Ward","起风了（翻自 高橋優） + 阿泱","起风了 + 买辣椒也用券","带你去旅行 + 亦勋/林子琪","LOSER (Live) - (失败者) + BIGBANG","삐딱하게 (Crooked) (Live) - (狂放) + G-DRAGON")


set wsh=createobject("wscript.shell")



wscript.Sleep 1000
wsh.popup "程序运行请回车继续.....","0","提示窗口","1"

Rem 使用for循环 动态操作剪贴板 实现输出 中文字符
for i = 0 to ubound(li)
wscript.sleep 200
Rem 复制到剪贴板
Clipboard="MsHta vbscript:ClipBoardData.setData(""Text"","""&li(i)&""")(Window.Close)"
wsh.Run(Clipboard)

Rem 粘贴到搜索框
wsh.popup "已复制到剪贴板,请准备粘贴.....","0","提示窗口","1"
wscript.sleep 1000
wsh.sendKeys "^v"
wscript.sleep 200
wsh.sendKeys "{enter}"

Rem 删除搜索框上一条内容
wsh.popup "下载完后请回车确认.....","0","提示窗口","1"
wscript.sleep 1000
wsh.sendKeys "^a"
wsh.sendKeys "{del}"

Rem wsh.popup li(i)




Next

wscript.quit

