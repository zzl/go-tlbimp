package typelib

import (
	"github.com/zzl/go-win32api/win32"
	"strings"
)

type ParamFlags struct {
	In       bool
	Out      bool
	Retval   bool
	Optional bool
}

func (me ParamFlags) String() string {
	var parts []string
	if me.In {
		parts = append(parts, "in")
	}
	if me.Out {
		parts = append(parts, "out")
	}
	if me.Retval {
		parts = append(parts, "retval")
	}
	if me.Optional {
		parts = append(parts, "optional")
	}
	return strings.Join(parts, ", ")
}

type ParamInfo struct {
	Name  string
	Type  *VarType
	Flags ParamFlags
}

func NewParamInfo(pTypeInfo *win32.ITypeInfo, pFuncDesc *win32.FUNCDESC,
	name string, pParamDesc win32.ELEMDESC, cParams int, index int) *ParamInfo {

	info := &ParamInfo{
		Name: name,
	}

	idlFlags := uint32(pParamDesc.IdldescVal().WIDLFlags)
	if idlFlags&win32.IDLFLAG_FIN != 0 {
		info.Flags.In = true
	}
	if idlFlags&win32.IDLFLAG_FOUT != 0 {
		info.Flags.Out = true
	}
	if idlFlags&win32.IDLFLAG_FRETVAL != 0 {
		info.Flags.Retval = true
	}
	if pFuncDesc.CParamsOpt == -1 && index == int(pFuncDesc.CParams-1) ||
		index >= cParams-int(pFuncDesc.CParamsOpt) {
		info.Flags.Optional = true
	}

	info.Type = NewVarType(pTypeInfo, &pParamDesc.Tdesc)
	return info
}
