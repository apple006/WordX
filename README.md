调用方法: `WordX.parse(file, entity, dataMap);`  
    file:Word源文件,只支持.docx格式  
    entity:要输出其属性的实体对象  
    dataMap：补充的自定义属性与值HashMap  
只能运行于Java8及以上版本  

### WordX根据Word中的书签来绑定数据的。书签的某些内容有特殊含义。  
`CompanyName` 将会输出对象的CompanyName属性值。  
`CompanyName_56` WordX会忽略最末尾的_数字，所以最终效果跟上面是一样的。  
`CompanyName_EW` 输出对象的CompanyName属性值。并且会在两端补空格，使跟书签选中的Word内容字符数相同。
`CompanyName_8W` 输出对象的CompanyName属性值。当长度超过8个英文字符时将截断。
`CompanyName_EW_8W_56` 多个修饰写在一起也是可以的。
`Accept_Check` 如果Accept属性值不等于0或空，书签选中的Word内容的方框字符“□”，将会打上勾。
`Sex_2_Radio` 如果Sex属性值等于中间的数字2，书签选中的Word内容的方框字符“□”，将会打上勾。
`Sex_00_Radio` 如果书签选中的Word内容包含Sex属性值，则书签选中的Word内容的方框字符“□”，将会打上勾。
`UpdatedAt_Fmt` 对UpdatedAt属性值格式化后输出。格式化字符串为书签选中的Word内容。
    格式化字符串一般按Java的来，如`%,.2f`、`%03d`，如果格式化字符串不包含%号，程序会在最前面补%号。
    日期的格式化按Golang的来，如`2006-01-02 15:04:05`。不同的数字代表不同的时间部份。
`Sex_Enum` 将Sex属性值转换为枚举字符串输出。同时需要有`SexEnum`属性值，其值应类似这样：`0未知 1男 2女`
`SignUrl_Img` 以SignUrl属性值指定的图片替换书签选中的Word图片。属性值可以是path、url、byte[]。
`X_Delete` 删除书签选中的Word内容。这里X不是属性名而是固定字符。
`X_Delete_Last` 仅当最后一次循环，书签选中的Word内容将被删除。
`Users_6List58_9Copy_Item` 遍历Users集合，并用Item代表一个元素。每个元素对应一份书签选中的Word内容。
    其中数字的含义：最少循环(复制)6次（即使没有6个元素），最多循环58次，从第9行（存在第0行）开始。
    这些数字参数都是可缺省的。
`Item_CompanyNam`e 输出遍历子元素Item的CompanyName属性值。
`Item_ListIndex` ListIndex比较特殊，输出的是遍历子元素Item的索引号，而不是属性值
