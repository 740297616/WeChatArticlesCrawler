> 本项目仅作个人学习使用，此分支由main分支延申，针对珠科公众号进行信息整理

#项目状态
项目目标：抓取文章阅读量、点赞量、在看量、标题、发布时间、访问链接、小尾巴(已分类)并生成Excel表格

已实现：阅读量、点赞量、在看量、标题、发布时间、访问链接、UI界面

待完成：1、**完全获取历史文章信息（因本人技术原因 + 平台限制暂无法实现）  2、在对小尾巴进行第一次筛选时根据有无“责编”判断是否小绿书，以此完善爬取
      

Notice:因某些特殊原因，暂不考虑实现自动登录WX获取敏感信息如key、pass_ticket、appmsg_token、cookie、user-agent等

#Usage
**Step1：**“微信公众平台”获取fakeid,token,cookie,user-agent等

**Step2：**"微信客户端 + Fiddler"访问目标公众号随意一篇文章，获取key,pass_ticket,appmsg_token,cookie,user-agent

[参考文章](https://blog.lydia0.cn/index.php/Python/66.html)

#实现思路
**思考**

1）如何批量获取文章链接？

2）文章中如何提取点赞、阅读、在看量？电脑上浏览文章只显示内容，没有阅读量、点赞量、评论

**分析**

1）Step1中需要提交data：

number表示从第number页开始爬取，为5的倍数，从0开始。如0、5、10……

token可以使用Chrome自带的工具进行获取

fakeid是公众号独一无二的一个id，等同于后面的__biz

2）Step2中的请求参数:

__biz对应公众号的信息，唯一

mid、sn、idx分别对应每篇文章的url的信息，需要从url中进行提取

key、appmsg_token从fiddler上复制即可

pass_ticket对应的文章的信息，貌似影响不大，也可以直接从fiddler复制

data中appmsg_type不可缺，否则无法获取like_num



**实现**

1）从微信公众平台批量获取文章链接，在此同时查看响应数据的时候发现了'title''create_time'即文章标题和发布时间戳

2）模拟微信客户端请求文章返回json，抓取阅读量read_num，在看量old_like_num,点赞量like_num

3）创建Excel表格，写入并保存
