rem 打字网----月桂女神
Dim wsh,a,li


li = Array("chuan shuo man chang hao han ru shi shi ban","ji zai zhe duan huang huang bu an","yan se jin huang a bo luo de guang mang","que bi bu shang da fu ni de yong gan","mei you yi zhong ai ke yi zai zi you zhi shang","da fu ni de shang hua shen yue gui shu jue jiang","yue gui shu piao xiang na ye feng lian ye guang wo de ai hen bu yi yang","su jing de lian shang cong bu mo nong zhuang jian chi zi ji xi huan","yue gui shu piao xiang yun chan rao xing guang wo yao you hua jiu jiang","wu bian de hai yang na liao kuo de xiang xiang bi shui dou bu ping fan","sen lin he pan a bo luo zai zhui gan","ku zhe dai shang da fu ni de gui guan","bei su fu de ai yi jing mei you le wen nuan","da fu ni de shang xin tong qian nian jian liu chuan","yue gui shu piao xiang na ye feng lian ye guang wo de ai hen bu yi yang","su jing de lian shang cong bu mo nong zhuang jian chi zi ji xi huan","yue gui shi piao xiang yun chan rao xing guang wo yao you hua jiu jiang","wu bian de hai yang na liao kuo de xiang xiang bi shui dou bu ping fan","ai yao huang ai kao an wo xing xiang qian fang xun zhao gui guan","yue gui shu piao xiang na ye feng lian ye guang wo de a hen bu yi yang","su jing de lian shang cong bu mo nong zhuang jian chi zi ji xi huan","yue gui sui piao xiang yun chan rao xing guang wo yao you hua jiu jiang","wu bian de hai yang na liao kuo de xiang xiang bi shui dou bu ping fan")


set wsh=createobject("wscript.shell")


wscript.Sleep 4000

for i = 0 to ubound(li)-1
wscript.sleep 2000
wsh.sendKeys li(i)
wscript.Sleep 1300
wsh.sendKeys " "
wscript.Sleep 140
wsh.sendKeys "{enter}{enter}"

Next

wscript.quit

