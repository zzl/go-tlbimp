package typelib

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-tlbimp/utils"
	"github.com/zzl/go-win32api/v2/win32"
	"strconv"
	"strings"
	"syscall"
	"unsafe"
)

type VarType struct {
	Name  string
	Size  int
	Align int

	Native        bool //numbers,uintptr
	Unsigned      bool
	Pointer       bool //*,unsafe.Pointer
	Array         bool //[]
	Struct        bool //struct,union
	Interface     bool //com
	DispInterface bool //com

	RefType *VarType

	PVarCastExpr string
}

func NewVarType(pTypeInfo *win32.ITypeInfo, pTypeDesc *win32.TYPEDESC) *VarType {
	return _newVarType(pTypeInfo, pTypeDesc, true)
}

func _newVarType(pTypeInfo *win32.ITypeInfo, pTypeDesc *win32.TYPEDESC,
	resolveIndirectRefType bool) *VarType {

	var t VarType
	switch win32.VARENUM(pTypeDesc.Vt) {
	case win32.VT_I2:
		t.Name = "int16"
		t.Size = 2
		t.Native = true
		t.PVarCastExpr = "$.IValVal()"
	case win32.VT_I4:
		t.Name = "int32"
		t.Size = 4
		t.Native = true
		t.PVarCastExpr = "$.LValVal()"
	case win32.VT_R4:
		t.Name = "float32"
		t.Size = 4
		t.Native = true
		t.PVarCastExpr = "$.FltValVal()"
	case win32.VT_R8:
		t.Name = "float64"
		t.Size = 8
		t.Native = true
		t.PVarCastExpr = "$.DblValVal()"
	case win32.VT_CY:
		t.Name = "win32.CY"
		t.Size = 8
		t.Struct = true
		t.PVarCastExpr = "$.CyValVal()"
	case win32.VT_DATE:
		t.Name = "ole.Date"
		t.Size = 8
		t.Native = true
		t.PVarCastExpr = "ole.Date($.DateVal())"
	case win32.VT_BSTR:
		t.Name = "win32.BSTR"
		t.Pointer = true
		t.Size = utils.PtrSize
		if resolveIndirectRefType {
			t.RefType = &VarType{
				Name:   "uint16",
				Native: true,
			}
		}
		t.PVarCastExpr = "$.BstrValVal()"
	case win32.VT_DISPATCH:
		t.Name = "*win32.IDispatch"
		t.Pointer = true
		t.Size = utils.PtrSize
		if resolveIndirectRefType {
			t.RefType = &VarType{
				Name:      "win32.IDispatch",
				Interface: true,
			}
		}
		t.PVarCastExpr = "$.PdispValVal()"
	case win32.VT_ERROR:
		t.Name = "win32.HRESULT"
		t.Size = 4
		t.Native = true
		t.PVarCastExpr = "$.ScodeVal()"
	case win32.VT_BOOL:
		t.Name = "win32.VARIANT_BOOL"
		t.Size = 2
		t.Native = true
		t.PVarCastExpr = "$.BoolValVal()"
	case win32.VT_VARIANT:
		t.Name = "win32.VARIANT"
		v := win32.VARIANT{}
		t.Size = int(unsafe.Sizeof(v))
		t.Align = int(unsafe.Alignof(v))
		t.Struct = true
		t.PVarCastExpr = "*$"
	case win32.VT_UNKNOWN:
		t.Name = "*win32.IUnknown"
		t.Size = utils.PtrSize
		t.Pointer = true
		if resolveIndirectRefType {
			t.RefType = &VarType{
				Name:      "win32.IUnknown",
				Interface: true,
			}
		}
		t.PVarCastExpr = "$.PunkValVal()"
	case win32.VT_DECIMAL:
		t.Name = "win32.DECIMAL"
		d := win32.DECIMAL{}
		t.Size = int(unsafe.Sizeof(d))
		t.Align = int(unsafe.Alignof(d))
		t.Struct = true
		t.PVarCastExpr = "$.DecValVal()"
	case win32.VT_I1:
		t.Name = "int8"
		t.Size = 1
		t.Native = true
		t.PVarCastExpr = "int8($.CValVal())"
	case win32.VT_UI1:
		t.Name = "byte"
		t.Size = 1
		t.Unsigned = true
		t.Native = true
		t.PVarCastExpr = "$.BValVal()"
	case win32.VT_UI2:
		t.Name = "uint16"
		t.Size = 2
		t.Unsigned = true
		t.Native = true
		t.PVarCastExpr = "$.UiValVal()"
	case win32.VT_UI4:
		t.Name = "uint32"
		t.Size = 4
		t.Unsigned = true
		t.Native = true
		t.PVarCastExpr = "$.UintValVal()"
	case win32.VT_I8:
		t.Name = "int64"
		t.Size = 8
		t.Native = true
		t.PVarCastExpr = "$.LlValVal()"
	case win32.VT_UI8:
		t.Name = "uint64"
		t.Size = 8
		t.Unsigned = true
		t.Native = true
		t.PVarCastExpr = "$.UllValVal()"
	case win32.VT_INT:
		t.Name = "int32"
		t.Size = 4
		t.Native = true
		t.PVarCastExpr = "$.LValVal()"
	case win32.VT_UINT:
		t.Name = "uint32"
		t.Size = 4
		t.Native = true
		t.PVarCastExpr = "$.UintValVal()"
	case win32.VT_VOID:
		t.Name = ""
		t.Size = 0
	case win32.VT_HRESULT:
		t.Name = "win32.HRESULT"
		t.Size = 4
		t.PVarCastExpr = "$.ScodeVal()"
	case win32.VT_PTR:
		if resolveIndirectRefType {
			t.RefType = NewVarType(pTypeInfo, pTypeDesc.LptdescVal())
			if t.RefType.Name == "" { //void
				t.Name = "unsafe.Pointer"
			} else if t.RefType.Name == "unsafe.Pointer" {
				t.Name = "unsafe.Pointer"
			} else {
				t.Name = "*" + t.RefType.Name
			}
		} else {
			t.Name = "unsafe.Pointer"
		}
		t.Pointer = true
		t.Size = utils.PtrSize
		//t.PVarCastExpr = "??"
	case win32.VT_CARRAY:
		t.RefType = _newVarType(pTypeInfo, &pTypeDesc.LpadescVal().TdescElem, resolveIndirectRefType)
		t.Array = true
		dimCount := int(pTypeDesc.LpadescVal().CDims)
		bounds := unsafe.Slice((*win32.SAFEARRAYBOUND)(
			unsafe.Pointer(&pTypeDesc.LpadescVal().Rgbounds)), dimCount)
		t.Name = ""
		totalElemCount := 0
		for n, b := range bounds {
			elemCount := int(b.CElements)
			if n == 0 {
				totalElemCount = elemCount
			} else {
				totalElemCount *= elemCount
			}
			t.Name += "[" + strconv.Itoa(elemCount) + "]"
		}
		t.Name += t.RefType.Name
		t.Size = totalElemCount * t.RefType.Size
		t.Align = t.RefType.Align
	case win32.VT_SAFEARRAY:
		t.Name = "*win32.SAFEARRAY"
		a := win32.SAFEARRAY{}
		t.Size = int(unsafe.Sizeof(a))
		t.Align = int(unsafe.Alignof(a))
		t.Struct = true
		t.PVarCastExpr = "$.ParrayVal()"
	case win32.VT_USERDEFINED:
		var ptiRef *win32.ITypeInfo
		hr := pTypeInfo.GetRefTypeInfo(pTypeDesc.HreftypeVal(), &ptiRef)
		defer ptiRef.Release()
		win32.ASSERT_SUCCEEDED(hr)

		var bs com.BStr
		hr = ptiRef.GetDocumentation(win32.MEMBERID_NIL, bs.PBSTR(), nil, nil, nil)
		win32.ASSERT_SUCCEEDED(hr)
		//
		t.Name = utils.CapName(bs.ToStringAndFree())

		if strings.HasPrefix(t.Name, "MIDL_IWinTypes") {
			t.Native = true
			t.Name = "uintptr"
			t.Unsigned = true
			t.Size = utils.PtrSize
			break
		}

		if strings.HasPrefix(t.Name, "Wire") { //?
			t.Name = "win32." + t.Name[4:]
			t.Native = true
			t.Unsigned = true
			t.Size = utils.PtrSize
			break
		}

		var ptaRef *win32.TYPEATTR
		ptiRef.GetTypeAttr(&ptaRef)
		defer ptiRef.ReleaseTypeAttr(ptaRef)

		switch ptaRef.Typekind {
		case win32.TKIND_ENUM:
			//t.Native = true
			var pVarDesc *win32.VARDESC
			ptiRef.GetVarDesc(0, &pVarDesc)
			t = *NewVarType(ptiRef, &pVarDesc.ElemdescVar.Tdesc)
		case win32.TKIND_RECORD:
			if t.Name == "GUID" {
				t.Name = "syscall.GUID"
				t.Struct = true
				g := syscall.GUID{}
				t.Size = int(unsafe.Sizeof(g))
				t.Align = int(unsafe.Alignof(g))
				break
			}
			t.Struct = true
			t.Size, t.Align = getStructSize(ptiRef, ptaRef)
		case win32.TKIND_COCLASS: //?
			t.Interface = true
			var pta *win32.TYPEATTR
			ptiRef.GetTypeAttr(&pta)
			for n := uint32(0); n < uint32(pta.CImplTypes); n++ {
				var implType win32.IMPLTYPEFLAGS
				hr = ptiRef.GetImplTypeFlags(n, &implType)
				win32.ASSERT_SUCCEEDED(hr)
				var hRefType win32.HREFTYPE
				hr = ptiRef.GetRefTypeOfImplType(n, &hRefType)
				win32.ASSERT_SUCCEEDED(hr)
				var ptiImpl *win32.ITypeInfo
				hr = ptiRef.GetRefTypeInfo(hRefType, &ptiImpl)
				win32.ASSERT_SUCCEEDED(hr)
				var ptaImpl *win32.TYPEATTR
				ptiImpl.GetTypeAttr(&ptaImpl)
				if implType&win32.IMPLTYPEFLAG_FDEFAULT != 0 &&
					implType&win32.IMPLTYPEFLAG_FSOURCE == 0 {
					t.DispInterface = ptaImpl.Typekind == win32.TKIND_DISPATCH
				}
				ptiImpl.ReleaseTypeAttr(ptaImpl)
				ptiImpl.Release()
			}
			ptiRef.ReleaseTypeAttr(pta)
		case win32.TKIND_INTERFACE:
			t.Interface = true
		case win32.TKIND_DISPATCH:
			t.Interface = true
			t.DispInterface = true
		case win32.TKIND_ALIAS:
			if t.Name == "GUID" {
				t.Name = "syscall.GUID"
				t.Struct = true
				g := syscall.GUID{}
				t.Size = int(unsafe.Sizeof(g))
				t.Align = int(unsafe.Alignof(g))
				break
			}
			name0 := t.Name
			t.RefType = _newVarType(ptiRef, &ptaRef.TdescAlias, resolveIndirectRefType)
			t = *t.RefType
			if t.Native {
				//
			} else {
				t.Name = name0
			}
		case win32.TKIND_UNION:
			t.Struct = true
			t.Size, t.Align = getUnionSize(ptiRef, ptaRef)
		}
	case win32.VT_LPSTR:
		t.Name = "win32.PSTR"
		t.Pointer = true
		t.Size = utils.PtrSize
	case win32.VT_LPWSTR:
		t.Name = "win32.PWSTR"
		t.Pointer = true
		t.Size = utils.PtrSize
	case win32.VT_INT_PTR:
		t.Name = "uintptr"
		t.Native = true
		t.Size = utils.PtrSize
	case win32.VT_UINT_PTR:
		t.Name = "uintptr"
		t.Native = true
		t.Size = utils.PtrSize
	default:
		panic("???")
	}
	if t.Align == 0 {
		t.Align = t.Size
	}
	return &t
}

func getStructSize(pti *win32.ITypeInfo, pta *win32.TYPEATTR) (int, int) {
	count := int(pta.CVars)
	fieldSizes := make([]utils.SizeInfo, count)

	for n := 0; n < count; n++ {
		var pvd *win32.VARDESC
		hr := pti.GetVarDesc(uint32(n), &pvd)
		win32.ASSERT_SUCCEEDED(hr)
		vt := _newVarType(pti, &pvd.ElemdescVar.Tdesc, false)
		fieldSizes[n] = utils.SizeInfo{
			vt.Size, vt.Align,
		}
		pti.ReleaseVarDesc(pvd)
	}
	size := utils.StructSize(fieldSizes...)
	return size.TotalSize, size.AlignSize
}

func getUnionSize(pti *win32.ITypeInfo, pta *win32.TYPEATTR) (int, int) {
	count := int(pta.CVars)
	var maxSize, maxAlign int
	for n := 0; n < count; n++ {
		var pvd *win32.VARDESC
		hr := pti.GetVarDesc(uint32(n), &pvd)
		win32.ASSERT_SUCCEEDED(hr)
		vt := _newVarType(pti, &pvd.ElemdescVar.Tdesc, false)
		if vt.Align == 0 {
			vt.Align = vt.Size
		}
		if vt.Size > maxSize {
			maxSize = vt.Size
		}
		if vt.Align > maxAlign {
			maxAlign = vt.Align
		}
		pti.ReleaseVarDesc(pvd)
	}
	return maxSize, maxAlign
}
