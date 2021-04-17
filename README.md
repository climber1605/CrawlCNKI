# CrawlCNKI
这是一个针对知网的爬虫，用于爬取知网上指定期刊从2012年到2020年所有发表文献的信息，包括篇名、作者、期刊名、发表时间、被引用量、被下载量。
爬虫基于python3中的selenium库编写。
使用方法：
1. pip3 install requirement.txt
2. python3 crawl.py start end
   start end为指定期刊的下标范围，为了便于在多台机器上分配任务而引入。


