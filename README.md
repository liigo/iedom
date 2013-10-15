iedom
=====
iedom，用于操作IE网页DOM的LUA之C模块。

###〇、基本概念

	Document：文档，对应一个HTML网页（或内嵌网页），是对COM接口IHTMLDocument2的封装。
	Collection：成员集合，用于容纳多个成员（Element），是对COM接口IHTMLElementCollection的封装。
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

###五、在C/C++中调用

使用本库的的LUA代码，只能被C/C++代码加载并执行，且加载之前必须把一个存有IHTMLDocument2*指针的lightuserdata值存入全局表，并命名为 "iedom$currentdoc"。另参见dom_currentdoc()。

要保证在lua代码中正确加载此C模块iedom，还需要事先（通过C/C++代码）设置package.cpath。

比较完整的调用代码示例如下：

	static void InvokeLua_ondoc(IHTMLDocument2* pDoc)
	{
		char luafile[MAX_PATH + 1], luacpath[MAX_PATH + 1];
		if(getDllPathFile(luafile, MAX_PATH, "ondoc.lua") && getDllPathFile(luacpath, MAX_PATH, "?.dll"))
		{
			//create new lua state
			lua_State* L = lua_newstate(luaAlloc, NULL);
			luaL_openlibs(L);

			//update package.cpath to ensure 'require' can find iedom.dll at right place
			lua_getglobal(L, "package");
			lua_pushstring(L, luacpath);
			lua_setfield(L, -2, "cpath");
			lua_pop(L, 1);

			//pass pointer IHTMLDocument2* pDoc into c module, through a named global var
			pDoc->AddRef();
			lua_pushlightuserdata(L, pDoc);
			lua_setglobal(L, "iedom$currentdoc");

			//a global msgbox() function only for test/debug
			lua_register(L, "msgbox", msgbox);

			//execute the lua file
			int r = luaL_dofile(L, luafile);
			if(r == 0) {
				//success
			} else {
				//error
	#ifdef _DEBUG
				const char* errmsg = lua_tostring(L, -1);
				printf("luaL_dofile() error: %s\n", errmsg);
				MessageBoxA(0, errmsg, "ondoc.lua error:", MB_OK|MB_ICONERROR);
	#endif
			}
			lua_close(L);
		} else {
			// fails to get the path of .lua
			// log error
		}
	}

	//return the full path of the file in this dll directory: DllDir\file
	static bool getDllPathFile(char* buffer, unsigned int buflen, const char* file)
	{
		HMODULE hThisDllModule = GetModuleHandleA("auditsink.dll");
		//if(GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, getDllPathFile, &hThisDllModule))
		if(hThisDllModule)
		{
			DWORD r = GetModuleFileNameA(hThisDllModule, buffer, buflen);
			if(r == 0 || r == buflen) return false;
			char* slash = strrchr(buffer, '\\');
			if(slash == NULL) return false;
			*(slash + 1) = '\0';
			if(file) {
				int n = strlen(file);
				if((slash + n + 1 >= buffer + buflen)) return false; // file\0
				memcpy(slash + 1, file, n + 1);
			}
			return true;
		} else {
			return false;
		}
	}

	static void* luaAlloc(void* ud, void* ptr, unsigned int osize, unsigned int nsize)
	{
		if(nsize == 0) {
			free(ptr);
			return NULL;
		}
		return realloc(ptr, nsize);
	}

	static int msgbox(lua_State* L)
	{
		const char* msg = lua_tostring(L, 1);
		MessageBoxA(0, msg, "Message", MB_OK);
		return 0;
	}

注1：因为本库仅仅是对COM对象 IHTMLDocument2* 及其相关接口的操作封装，而这些COM对象指针往往仅存在于C/C++程序中，所以限制本库只能被C/C++调用并无不妥。

注2：在本库内部实现中，借助table对象的 \__gc 元方法做到了在Lua对象document/element/collecton被回收时调用相应COM对象的Release()方法。这意味着：1、IHTMLDocument2*传入本库前必须先AddRef()；2、本库必须在Lua 5.2或更高版本环境下执行，否则，因为Lua 5.1及以下版本中table没有\__gc元方法将导致COM对象无法正常释放。

###六、License
Opensource under the [MIT license](http://opensource.org/licenses/MIT).
Copyright (C) 2013 Liigo Zhuang.
