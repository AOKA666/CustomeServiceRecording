1.打开企微，同时按`shift`+`ctrl`+`alt`+`D`，进入调试模式
![调试模式](guide_images/1.png)  

2.打开数据录入界面，右键空白处，点击 ShowDevTools
![DevTools](guide_images/2.png)  

3.发现打开一个小窗口，选中 Network
![小窗口](guide_images/3.png)  

4.此时再点一下“数据录入”按钮，切回小窗口看一下，发现多了很多东西
![点击](guide_images/4.png)  

5.在小窗口里找一个带橙色图标的行，点一下，发现跳出详情
![请求](guide_images/5.png)  
![详情](guide_images/6.png)  

6.依次找到`Request Headers`,`Authentication`，旁边会有一个 Basic 开头的字符串，这就是我们要找的，复制下来粘贴即可
![找到](guide_images/7.png)  

7. **复制时注意不要多复制空格回车**
