package typelib

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/win32"
	"unsafe"
)

type FuncFlags struct {
	PropGet    bool
	PropPut    bool
	PropPutRef bool
	Restricted bool
	Hidden     bool
	Vararg     bool
}

type FuncInfo struct {
	Id         win32.MEMBERID
	Name       string
	Doc        string
	Flags      FuncFlags
	Params     []*ParamInfo
	ReturnType *VarType
}

func NewFuncInfo(pTypeInfo *win32.ITypeInfo, pTypeAttr *win32.TYPEATTR,
	pFuncDesc *win32.FUNCDESC, dispFunc bool) *FuncInfo {

	var hr win32.HRESULT
	if pTypeAttr == nil {
		hr = pTypeInfo.GetTypeAttr(&pTypeAttr)
		win32.ASSERT_SUCCEEDED(hr)
		defer pTypeInfo.ReleaseTypeAttr(pTypeAttr)
	}

	info := &FuncInfo{}

	if pTypeAttr.Typekind == win32.TKIND_DISPATCH {
		info.Id = pFuncDesc.Memid
	}

	info.Flags.PropGet = pFuncDesc.Invkind&win32.INVOKE_PROPERTYGET != 0
	info.Flags.PropPut = pFuncDesc.Invkind&win32.INVOKE_PROPERTYPUT != 0
	info.Flags.PropPutRef = pFuncDesc.Invkind&win32.INVOKE_PROPERTYPUTREF != 0
	if pFuncDesc.WFuncFlags&uint16(win32.FUNCFLAG_FRESTRICTED) != 0 {
		info.Flags.Restricted = true
	}
	if pFuncDesc.WFuncFlags&uint16(win32.FUNCFLAG_FHIDDEN) != 0 {
		info.Flags.Hidden = true
	}
	if pFuncDesc.CParamsOpt == -1 {
		info.Flags.Vararg = true
	}

	var bsName, bsDoc com.BStr
	hr = pTypeInfo.GetDocumentation(pFuncDesc.Memid, bsName.PBSTR(), bsDoc.PBSTR(), nil, nil)
	win32.ASSERT_SUCCEEDED(hr)

	info.Name = bsName.ToStringAndFree()
	info.Doc = bsDoc.ToStringAndFree()

	//
	const maxNames = 64
	var bsNames [maxNames]win32.BSTR
	var cNames uint32
	pTypeInfo.GetNames(pFuncDesc.Memid, &bsNames[0], maxNames, &cNames)
	if cNames < uint32(pFuncDesc.CParams+1) {
		bsNames[cNames] = win32.SysAllocString(win32.StrToPwstr("rhs"))
	}
	defer func() {
		for n := 0; n < maxNames; n++ {
			//bsNames[n].Free()
			if bsNames[n] != nil {
				win32.SysFreeString(bsNames[n])
			}
		}
	}()

	elemDescParams := unsafe.Slice(pFuncDesc.LprgelemdescParam, pFuncDesc.CParams)
	cParams := int(pFuncDesc.CParams)
	if dispFunc {
		for n := 0; n < cParams; n++ {
			pParamDesc := elemDescParams[n]
			idlFlags := uint32(pParamDesc.IdldescVal().WIDLFlags)
			if idlFlags&win32.IDLFLAG_FLCID != 0 || idlFlags&win32.IDLFLAG_FRETVAL != 0 {
				cParams = n
				break
			}
		}
	}

	for n := 0; n < cParams; n++ {
		pParamDesc := elemDescParams[n]
		name := win32.BstrToStr(bsNames[n+1])
		param := NewParamInfo(pTypeInfo, pFuncDesc, name, pParamDesc, cParams, n)
		info.Params = append(info.Params, param)
	}

	info.ReturnType = NewVarType(pTypeInfo, &pFuncDesc.ElemdescFunc.Tdesc)
	return info
}
