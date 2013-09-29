#include "iedom.h"

extern "C" {
	#include <lua.h>
}

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;

#include <atlcom.h>
#include <mshtml.h>
#include <exdispid.h>
#include <exdisp.h>
#include <comutil.h>

#include <assert.h>


// 使用本库的的LUA代码，只能被C/C++代码加载并执行，且加载之前必须把一个table t存入全局表，
// 命名为 "iedom$currentdoc"，并令 t[0] = IHTMLDocument* 。参见dom_currentdoc()。
// TODO: 增加调用示例代码

struct LuaFunc {
	const char* name;
	int (*func) (lua_State*);
};

// create_lua_funcs_table() push a new created table on stack top, which contains all functions.
// if tname is no NULL, will save the table to Registry, to be take back later by using get_lua_funcs_table(L,tname).
// it is recommended to use a prefix for tname to avoid name confusion with other names in global Registry.
// by liigo, 20130906.
static void create_lua_funcs_table(lua_State* L, LuaFunc* funcs, const char* tname) {
	lua_newtable(L);
	int t = lua_gettop(L);
	while(funcs && funcs->name && funcs->func) {
		lua_pushcfunction(L, funcs->func);
		lua_setfield(L, t, funcs->name);
		funcs++;
	}
	if(tname) {
		lua_pushvalue(L, -1);
		lua_setfield(L, LUA_REGISTRYINDEX, tname);
	}
}

// push the table on stack top
// returns nil if no prev create_lua_funcs_table() called with the same name.
static void get_lua_funcs_table(lua_State* L, const char* tname) {
	lua_getfield(L, LUA_REGISTRYINDEX, tname);
}

// arg: table
// keep the table on stack top after return
static void add_table_funcs(lua_State* L, LuaFunc* funcs) {
	int t = lua_gettop(L);
	while(funcs && funcs->name && funcs->func) {
		lua_pushcfunction(L, funcs->func);
		lua_setfield(L, t, funcs->name);
		funcs++;
	}
}

static void lua_print(lua_State* L, const char* msg) {
	lua_getglobal(L, "print");
	lua_pushstring(L, msg);
	lua_call(L, 1, 0);
}

static void report_lua_error(lua_State* L, const char* errmsg) {
	lua_pushstring(L, errmsg);
	lua_error(L);
}

// arg: doc, ...
// return IHTMLDocument2*
static IHTMLDocument2* getdoc(lua_State* L) {
	if(!lua_istable(L, 1))
		report_lua_error(L, "require doc value as first parameter");
	lua_rawgeti(L, 1, 0);
	IHTMLDocument2* pDoc = (IHTMLDocument2*) lua_touserdata(L, -1);
	lua_pop(L, 1);
	if(pDoc == NULL)
		report_lua_error(L, "invalid doc value");
	return pDoc;
}

// no arg, return current document
static int dom_currentdoc(lua_State* L) {
	lua_getglobal(L, "iedom$currentdoc");
	if(lua_isnil(L, -1)) {
		report_lua_error(L, "[iedom] No current document value available.");
	}
	return 1;
}

// arg: collection, ...
// return IHTMLElementCollection* (not NULL)
static IHTMLElementCollection* getcollection(lua_State* L) {
	if(!lua_istable(L, 1))
		report_lua_error(L, "require collection value as first parameter");
	lua_rawgeti(L, 1, 0);
	IHTMLElementCollection* c = (IHTMLElementCollection*) lua_touserdata(L, -1);
	lua_pop(L, 1);
	if(c == NULL)
		report_lua_error(L, "invalid collection value");
	return c;
}

// arg: element, ...
// return IHTMLElement* (not NULL)
static IHTMLElement* getelement(lua_State* L) {
	if(!lua_istable(L, 1))
		report_lua_error(L, "require element value as first parameter");
	lua_rawgeti(L, 1, 0);
	IHTMLElement* e = (IHTMLElement*) lua_touserdata(L, -1);
	lua_pop(L, 1);
	if(e == NULL)
		report_lua_error(L, "invalid element value");
	return e;
}

static void new_collection(lua_State* L, IHTMLElementCollection* collection);

// get all children element collection
// no arg, return collection
static int element_all(lua_State* L) {
	IDispatch* dispCollection = NULL;
	IHTMLElementCollection* pCollection = NULL;
	IHTMLElement* pElement = getelement(L);
	if(pElement->get_all(&dispCollection) == S_OK && 
			dispCollection->QueryInterface(IID_IHTMLElementCollection, (void**) &pCollection) == S_OK)
	{
		new_collection(L, pCollection);
		return 1;
	}
	lua_pushnil(L);
	return 1;
}

// get or set attrubite
// get arg: element, attrname(string), flag(int,optional)
// set arg: element, attrname(string), attrvalue(sting or nil), flag(int,optional)
// if set attrvalue as nil, the attr with the name attrname will be removed.
// return attribute value string, or "" if error occurs. when set attrubite, return boolean.
static int element_attr(lua_State* L) {
	IHTMLElement* pElement = getelement(L);
	int top = lua_gettop(L);
	int flag = 0;
	if(lua_isnumber(L, top)) {
		flag = lua_tointeger(L, top); // flag: 0 1 2 4
	}
	_bstr_t sName = lua_tostring(L, 2);
	if(lua_isstring(L, 3) || lua_isnil(L, 3)) {
		// set attr
		if(lua_isstring(L, 3)) {
			_variant_t vValue = lua_tostring(L, 3);
			HRESULT r = pElement->setAttribute((BSTR)sName, (VARIANT)vValue, flag);
			lua_pushboolean(L, r == S_OK);
		} else {
			VARIANT_BOOL vOK;
			pElement->removeAttribute((BSTR)sName, flag, &vOK); // flag: 0 1
			lua_pushboolean(L, vOK == VARIANT_TRUE);
		}
		return 1;
	} else {
		// get attr
		VARIANT vValue;
		if(pElement->getAttribute((BSTR)sName, flag, &vValue) == S_OK && vValue.vt == VT_BSTR) {
			_bstr_t bstr = vValue.bstrVal;
			lua_pushstring(L, (char*)bstr);
		} else {
			lua_pushstring(L, "");
		}
		return 1;
	}
}

// get immediate children element collection
// no arg, return collection
static int element_children(lua_State* L) {
	IDispatch* dispCollection = NULL;
	IHTMLElementCollection* pCollection = NULL;
	IHTMLElement* pElement = getelement(L);
	if(pElement->get_children(&dispCollection) == S_OK && 
			dispCollection->QueryInterface(IID_IHTMLElementCollection, (void**) &pCollection) == S_OK) {
		new_collection(L, pCollection);
		return 1;
	}
	lua_pushnil(L);
	return 1;
}

// get or set class name
// arg: element, newclass(string,optional)
// retrun class string, or "" if error occurs. when set class, return boolean.
static int element_class(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_className(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_className((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}

// get or set id
// arg: element, newid(string,optional)
// retrun id string, or "" if error occurs. when set id, return boolean.
static int element_id(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_id(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_id((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}

// get or set innerhtml
// arg: element, newvalue(string,optional)
// retrun innerhtml string, or "" if error occurs. when set innerhtml, return boolean.
static int element_innerhtml(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_innerHTML(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_innerHTML((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}
static int element_innertext(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_innerText(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_innerText((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}
static int element_outerhtml(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_outerHTML(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_outerHTML((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}
static int element_outertext(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_outerText(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_outerText((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}

// get offsets: left, top, width, height
// no arg, return 4 offset integer values
static int element_offsets(lua_State* L) {
	IHTMLElement* pElement = getelement(L);
	long left, top, width, height;
	if(pElement->get_offsetLeft(&left) == S_OK && pElement->get_offsetTop(&top) == S_OK &&
		pElement->get_offsetWidth(&width) == S_OK && pElement->get_offsetHeight(&height) == S_OK)
	{
		lua_pushinteger(L, left); lua_pushinteger(L, top); lua_pushinteger(L, width); lua_pushinteger(L, height);
	} else {
		lua_pushnil(L); lua_pushnil(L); lua_pushnil(L); lua_pushnil(L);
	}
	return 4;
}

static void new_element(lua_State* L, IHTMLElement* element);

// return parent (element)
static int element_parent(lua_State* L) {
	IHTMLElement* pElement = getelement(L);
	IHTMLElement* pParent = NULL;
	if(pElement->get_parentElement(&pParent) == S_OK && pParent) {
		new_element(L, pParent);
		return 1;
	}
	lua_pushnil(L);
	return 1;
}

// return tag name (string)
static int element_tag(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(pElement->get_tagName(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}

// get or set title
// arg: element, newtitle(string,optional)
// retrun title string, or "" if error occurs. when set title, return boolean.
static int element_title(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLElement* pElement = getelement(L);
	if(lua_gettop(L) == 1 && pElement->get_title(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = pElement->put_title((BSTR)bstr);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}
	lua_pushstring(L, "");
	return 1;
}

static void new_element(lua_State* L, IHTMLElement* element) {
	LuaFunc element_funcs[] = {
		{ "all", element_all },
		{ "attr", element_attr },
		{ "children", element_children },
		{ "class", element_class },
		{ "id", element_id },
		{ "innerhtml", element_innerhtml },
		{ "innertext", element_innertext },
		{ "outerhtml", element_outerhtml },
		{ "outertext", element_outertext },
		{ "offsets", element_offsets },
		{ "parent", element_parent },
		{ "tag", element_tag },
		{ "title", element_title },
		{ NULL, NULL },
	};
	create_lua_funcs_table(L, element_funcs, NULL);
	lua_pushlightuserdata(L, element);
	lua_rawseti(L, -2, 0);
}

static int collection_len(lua_State* L) {
	IHTMLElementCollection* collection = getcollection(L);
	long len = 0;
	collection->get_length(&len);
	lua_pushinteger(L, len);
	return 1;
}

static int collection_item(lua_State* L);

static void new_collection(lua_State* L, IHTMLElementCollection* collection) {
	assert(collection);
	LuaFunc collection_funcs[] = {
		{ "len",  collection_len },
		{ "item", collection_item },
	//  { "tags", collection_tags },
		{ NULL, NULL },
	};
	create_lua_funcs_table(L, collection_funcs, NULL);
	lua_pushlightuserdata(L, collection);
	lua_rawseti(L, -2, 0); // collection[0] = IHTMLElementCollection*
}

// arg: collection (self), name (int or string), index (int,optional)
static int collection_item(lua_State* L) {
	IHTMLElementCollection* collection = getcollection(L);

	int argcount = lua_gettop(L);
	if(argcount < 2 || argcount > 3) {
		lua_pushnil(L);
		return 1;
	}

	VARIANT vName; VariantInit(&vName);
	CComBSTR bstr;
	if(lua_isnumber(L, 2)) {
		vName.vt = VT_I4;
		vName.lVal = lua_tointeger(L, 2) - 1; // 1 based
		long len = 0; collection->get_length(&len);
		if(vName.lVal < 0 || vName.lVal > len - 1) {
			lua_pushnil(L);
			return 1;
		}
	} else if(lua_isstring(L, 2)) {
		vName.vt = VT_BSTR;
		bstr = lua_tostring(L, 2);
		vName.bstrVal = bstr;
	} else {
		report_lua_error(L, "item() requires a integer or string value as its second parameter.");
	}

	VARIANT vIndex; VariantInit(&vIndex);
	if(argcount == 3) {
		vIndex.vt = VT_I4;
		vIndex.lVal = lua_tointeger(L, 3) - 1; // 1 based, valid?
	}

	IDispatch* pdisp = NULL;
	HRESULT r = collection->item(vName, vIndex, &pdisp); //invoke IHTMLElementCollection->item()
	if(r != S_OK || !pdisp) {
		lua_pushnil(L); return 1;
	}

	IHTMLElement* pElement = NULL;
	r = pdisp->QueryInterface(IID_IHTMLElement, (void**) &pElement);
	if(r == S_OK) {
		new_element(L, pElement);
		return 1;
	} else {
		IHTMLElementCollection* pCollection = NULL;
		r = pdisp->QueryInterface(IID_IHTMLElementCollection, (void**) &pCollection);
		if(r == S_OK) {
			new_collection(L, pCollection);
			return 1;
		}
	}

	lua_pushnil(L);
	return 1;
}

// arg: doc (self), 
static int doc_all(lua_State* L) {
	IHTMLDocument2* doc = getdoc(L);
	IHTMLElementCollection* collection = NULL;
	HRESULT r = doc->get_all(&collection);
	if(r != S_OK || collection == NULL) {
		report_lua_error(L, "[iedom] Invoke IHTMLDocument2->get_all() fails.");
	}
	new_collection(L, collection);
	return 1;
}

// get or set doc's url
// arg: doc, newurl(string,optional)
// retrun url string, or "" if error occurs. when set url, return boolean.
static int doc_url(lua_State* L) {
	BSTR bstr = NULL;
	IHTMLDocument2* doc = getdoc(L);
	if(lua_gettop(L) == 1 && doc->get_URL(&bstr) == S_OK) {
		_bstr_t s = bstr;
		lua_pushstring(L, (char*)s);
		return 1;
	} else if(lua_gettop(L) == 2 && lua_isstring(L, 2)) {
		_bstr_t s = lua_tostring(L, 2);
		HRESULT r = doc->put_URL((BSTR)s);
		lua_pushboolean(L, r == S_OK);
		return 1;
	}

	lua_pushstring(L, "");
	return 1;
}

int luaopen_iedom(lua_State* L) {
	LuaFunc dom_funcs[] = {
		//{ "currentdoc", dom_currentdoc },
		{ NULL, NULL },
	};
	create_lua_funcs_table(L, dom_funcs, NULL); // dom to be return

	LuaFunc doc_funcs[] = {
		{ "all", doc_all },
		{ "url", doc_url },
		{ NULL, NULL },
	};
	dom_currentdoc(L);
	add_table_funcs(L, doc_funcs);

	lua_setfield(L, -2, "currentdoc"); // dom.currentdoc = current documnet
	return 1;
}
