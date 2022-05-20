package typelib

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/win32"
)

type FieldInfo struct {
	Name  string
	Doc   string
	Type  *VarType
	Value interface{}
}

func NewFieldInfo(pTypeInfo *win32.ITypeInfo, pVarDesc *win32.VARDESC, withValue bool) *FieldInfo {
	fi := &FieldInfo{}
	var bsName com.BStr
	var cNames uint32
	hr := pTypeInfo.GetNames(pVarDesc.Memid, bsName.PBSTR(), 1, &cNames)
	win32.ASSERT_SUCCEEDED(hr)

	fi.Name = bsName.ToStringAndFree()

	var bsDoc com.BStr
	hr = pTypeInfo.GetDocumentation(pVarDesc.Memid, nil, bsDoc.PBSTR(), nil, nil)
	win32.ASSERT_SUCCEEDED(hr)
	fi.Doc = bsDoc.ToStringAndFree()

	fi.Type = NewVarType(pTypeInfo, &pVarDesc.ElemdescVar.Tdesc)

	if withValue {
		fi.Value = (*ole.Variant)(pVarDesc.LpvarValueVal()).Value()
	}
	return fi
}
