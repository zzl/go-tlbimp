package typelib

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
	"syscall"
)

type TypeFlags struct {
	Default       bool
	Hidden        bool
	Dual          bool
	OleAutomation bool
	Restricted    bool
}

type ImplType struct {
	Name          string
	Guid          syscall.GUID
	Default       bool
	Source        bool
	DispInterface bool
}

type TypeInfo struct {
	Name string
	Doc  string

	Guid syscall.GUID
	Kind win32.TYPEKIND

	FuncCount  int
	FieldCount int

	Flags   TypeFlags
	RelType *VarType //for alias

	Fields []*FieldInfo
	Funcs  []*FuncInfo

	//for interface
	Super *TypeInfo

	//
	DualInterface *TypeInfo
	DispInterface bool

	//for coclass
	ImplTypes []*ImplType

	Size, Align int
}

func NewTypeInfo(p *win32.ITypeInfo) *TypeInfo {
	info := &TypeInfo{}

	var bsName, bsDoc com.BStr
	hr := p.GetDocumentation(win32.MEMBERID_NIL, bsName.PBSTR(), bsDoc.PBSTR(), nil, nil)
	win32.ASSERT_SUCCEEDED(hr)

	info.Name = bsName.ToStringAndFree()
	info.Doc = bsDoc.ToStringAndFree()

	var pAttr *win32.TYPEATTR
	hr = p.GetTypeAttr(&pAttr)
	win32.ASSERT_SUCCEEDED(hr)

	info.Guid = pAttr.Guid
	info.Kind = pAttr.Typekind
	info.FuncCount = int(pAttr.CFuncs)
	info.FieldCount = int(pAttr.CVars)

	if pAttr.WTypeFlags&uint16(win32.TYPEFLAG_FHIDDEN) != 0 {
		info.Flags.Hidden = true
	}
	if pAttr.WTypeFlags&uint16(win32.TYPEFLAG_FDUAL) != 0 {
		info.Flags.Dual = true
	}
	if pAttr.WTypeFlags&uint16(win32.TYPEFLAG_FOLEAUTOMATION) != 0 {
		info.Flags.OleAutomation = true
	}
	if pAttr.WTypeFlags&uint16(win32.TYPEFLAG_FRESTRICTED) != 0 {
		info.Flags.Restricted = true
	}

	if info.Kind == win32.TKIND_ALIAS {
		info.RelType = NewVarType(p, &pAttr.TdescAlias)
	}

	defer p.ReleaseTypeAttr(pAttr)

	//
	if info.Kind == win32.TKIND_ENUM {
		for n := 0; n < info.FieldCount; n++ {
			var pVarDesc *win32.VARDESC
			hr = p.GetVarDesc(uint32(n), &pVarDesc)
			win32.ASSERT_SUCCEEDED(hr)

			field := NewFieldInfo(p, pVarDesc, true)
			p.ReleaseVarDesc(pVarDesc)

			info.Fields = append(info.Fields, field)
		}
	} else if info.Kind == win32.TKIND_RECORD {
		for n := 0; n < info.FieldCount; n++ {
			var pVarDesc *win32.VARDESC
			hr = p.GetVarDesc(uint32(n), &pVarDesc)
			win32.ASSERT_SUCCEEDED(hr)

			field := NewFieldInfo(p, pVarDesc, false)
			p.ReleaseVarDesc(pVarDesc)

			info.Fields = append(info.Fields, field)
		}
		info.Size, info.Align = getStructSize(p, pAttr)
	} else if info.Kind == win32.TKIND_UNION {
		for n := 0; n < info.FieldCount; n++ {
			var pVarDesc *win32.VARDESC
			hr = p.GetVarDesc(uint32(n), &pVarDesc)
			win32.ASSERT_SUCCEEDED(hr)

			field := NewFieldInfo(p, pVarDesc, false)
			p.ReleaseVarDesc(pVarDesc)

			info.Fields = append(info.Fields, field)
		}
		info.Size, info.Align = getUnionSize(p, pAttr)
	}

	if info.Kind == win32.TKIND_INTERFACE {
		if pAttr.CImplTypes > 0 {
			var hRefType win32.HREFTYPE
			hr = p.GetRefTypeOfImplType(0, &hRefType)
			win32.ASSERT_SUCCEEDED(hr)

			var ptiImpl *win32.ITypeInfo
			hr = p.GetRefTypeInfo(hRefType, &ptiImpl)
			win32.ASSERT_SUCCEEDED(hr)

			info.Super = NewTypeInfo(ptiImpl)
			ptiImpl.Release()
		} else {
			//iunknown?
		}
		for n := 0; n < info.FuncCount; n++ {
			var pFuncDesc *win32.FUNCDESC
			hr = p.GetFuncDesc(uint32(n), &pFuncDesc)
			fi := NewFuncInfo(p, pAttr, pFuncDesc, false)
			p.ReleaseFuncDesc(pFuncDesc)
			info.Funcs = append(info.Funcs, fi)
		}
	}

	if info.Kind == win32.TKIND_DISPATCH {
		if pAttr.CImplTypes > 0 {
			var hRefType win32.HREFTYPE
			hr = p.GetRefTypeOfImplType(0, &hRefType)
			win32.ASSERT_SUCCEEDED(hr)

			var ptiImpl *win32.ITypeInfo
			hr = p.GetRefTypeInfo(hRefType, &ptiImpl)
			win32.ASSERT_SUCCEEDED(hr)

			info.Super = NewTypeInfo(ptiImpl)
			ptiImpl.Release()
		} else {
			//iunknown?
		}

		info.DispInterface = true
		if info.Flags.Dual {
			var refType uint32
			hr := p.GetRefTypeOfImplType(^uint32(0), &refType)
			win32.ASSERT_SUCCEEDED(hr)

			var pti *win32.ITypeInfo
			hr = p.GetRefTypeInfo(refType, &pti)
			win32.ASSERT_SUCCEEDED(hr)

			info.DualInterface = NewTypeInfo(pti)
		}

		for n := 0; n < info.FuncCount; n++ {
			var pFuncDesc *win32.FUNCDESC
			hr = p.GetFuncDesc(uint32(n), &pFuncDesc)
			fi := NewFuncInfo(p, pAttr, pFuncDesc, true)
			p.ReleaseFuncDesc(pFuncDesc)
			info.Funcs = append(info.Funcs, fi)
		}
	}

	//
	if info.Kind == win32.TKIND_COCLASS {

		var implType win32.IMPLTYPEFLAGS
		var hRefType win32.HREFTYPE
		var ptiImpl *win32.ITypeInfo
		var bsName com.BStr
		var pImplAttr *win32.TYPEATTR

		for n := uint32(0); n < uint32(pAttr.CImplTypes); n++ {
			hr = p.GetImplTypeFlags(n, &implType)
			if win32.FAILED(hr) {
				break
			}
			hr = p.GetRefTypeOfImplType(n, &hRefType)
			if win32.FAILED(hr) {
				break
			}
			win32.ASSERT_SUCCEEDED(hr)

			hr = p.GetRefTypeInfo(hRefType, &ptiImpl)
			win32.ASSERT_SUCCEEDED(hr)

			ptiImpl.GetDocumentation(win32.MEMBERID_NIL, bsName.PBSTR(), nil, nil, nil)
			ptiImpl.GetTypeAttr(&pImplAttr)

			intf := &ImplType{
				Name:          bsName.ToStringAndFree(),
				Guid:          pImplAttr.Guid,
				Default:       implType&win32.IMPLTYPEFLAG_FDEFAULT != 0,
				Source:        implType&win32.IMPLTYPEFLAG_FSOURCE != 0,
				DispInterface: pImplAttr.Typekind == win32.TKIND_DISPATCH,
			}

			info.ImplTypes = append(info.ImplTypes, intf)
			ptiImpl.ReleaseTypeAttr(pImplAttr)
			ptiImpl.Release()
		}
	}
	return info
}

func (this *TypeInfo) GetField(index int) *FieldInfo {
	return this.Fields[index]
}

func (this *TypeInfo) GetFunc(index int) *FuncInfo {
	return this.Funcs[index]
}
