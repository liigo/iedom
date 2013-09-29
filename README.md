iedom
=====
iedom，用于操作IE网页DOM的LUA之C模块。

###〇、基本概念

	Document：文档，对应一个HTML网页（或内嵌网页），是对COM接口IHTMLDocument2的封装。
	Collection：成员集合，用于容纳哦多个成员（Element），是对COM接口IHTMLElementCollection的封装。
	Element：成员，对应HTML文本内的一个标签（Tag），是对COM接口IHTMLElement或其子接口的封装。

这些概念与微软网页DOM中的概念是一一对应的。更多更详细资料请参考MSDN：

	IHTMLDocument2 http://msdn.microsoft.com/en-us/library/aa752574(v=vs.85).aspx
	IHTMLElement   http://msdn.microsoft.com/en-us/library/aa752279(v=vs.85).aspx
	IHTMLElementCollection http://msdn.microsoft.com/en-us/library/aa703928(v=vs.85).aspx

###一、加载iedom库，获取并操作当前Document文档

	local dom = require "iedom" --加载iedom模块
	local doc = dom.currentdoc  --取当前文档对象

文档（Document）对象有以下方法（第一个参数必须是对象自身）：

	doc:all()        --返回本文档所有成员的集合（Collection）
	doc:url([value]) --取URL或置URL文本。value为新的URL文本，仅当置URL时生效，否则可省略。取URL时返回文本，置URL时返回逻辑值。

###二、操作成员集合（Collection）

成员集合（Collection）对象有以下方法（第一个参数必须是对象自身）：

	c:len()       --取成员集合中的成员个数
	c:item(index) --取指定索引处的成员（Element），index取值范围是1到len()，index时无效返回nil。

###三、操作成员（Element）

成员（Element）对象有以下方法（第一个参数必须是对象自身）

	e:all()                  --返回所有子孙成员集合，出错时返回nil。介于本成员HTML开始标签和结束标签之间的所有成员，都在返回值集合内。
	e:attr(name,flag)        --取属性值，返回文本。name为属性名称文本，flag可为数值0,1,2,4，默认为0。没有指定名称的情况返回空文本""。
	e:attr(name,value,flag)  --置属性值，返回逻辑值表示设置是否成功。name和value均为文本类型，flag同上。如果value为nil，将删除该属性。
	e:children()             --返回直接子成员集合，出错时返回nil。相比all()，children()只返回包含儿子成员的集合，而不包含孙子成员。
	e:class([value])         --取类名或置类名文本（即HTML标签内的class属性值）。value为新类名文本，仅置类名时有效，否则可省略。取类名时返回文本，置类名时返回逻辑值。
	e:id([value])            --取ID或置ID文本（即HTML标签内的id属性值）。value为新ID文本，仅置ID时有效，否则可省略。取ID时返回文本，置ID时返回逻辑值。
	e:innerhtml([value])     --取和置 innerhtml,innertext,outerhtml,outertext。参数和返回值同上。下同。
	e:innertext([value])
	e:outerhtml([value])
	e:outertext([value])
	e:offsets()              --取偏移，返回4个整数值分别表示左/上/宽/高，出错时返回4个nil。
	e:parent()               --取父成员（Element），失败返回nil。
	e:tag()                  --取标签文本，返回值文本为全大写文本，如"BODY","DIV","P","SCRIPT"。
	e:title([value])         --取Title或置Title文本（Title通常用作ToolTip）。value为新Title文本，仅置Title时有效，否则可省略。取Title时返回文本，置ID时返回逻辑值。

###四、应用示例

	local dom = require "iedom"
	local doc = dom.currentdoc
	local all = doc:all()

	local i, len = all:len()
	for i=1, len, 1 do
		local element = all:item(i) 
		print(i .. " tag: " .. element:tag())
		if(element:tag() == "SPAN") then
			print("attr id = " .. element:attr("id"))	
		end
	end

###五、License
Opensource under the [MIT license](http://opensource.org/licenses/MIT).
Copyright (C) 2013 Liigo Zhuang.
