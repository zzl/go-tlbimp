package typelib

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-win32api/v2/win32"
)

type TypeLib struct {
	p *win32.ITypeLib
}

func NewTypeLibFromFile(filePath string) (*TypeLib, error) {
	var p *win32.ITypeLib
	hr := win32.LoadTypeLib(win32.StrToPwstr(filePath), &p)
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewTypeLib(p), nil
}

func NewTypeLib(p *win32.ITypeLib) *TypeLib {
	return &TypeLib{p: p}
}

func (this *TypeLib) Dispose() {
	if this.p == nil {
		return
	}
	this.p.Release()
}

func (this *TypeLib) GetName() string {
	var bs com.BStr
	this.p.GetDocumentation(win32.MEMBERID_NIL, bs.PBSTR(), nil, nil, nil)
	return bs.ToStringAndFree()
}

func (this *TypeLib) GetDoc() string {
	var bs com.BStr
	this.p.GetDocumentation(win32.MEMBERID_NIL, bs.PBSTR(), nil, nil, nil)
	return bs.ToStringAndFree()
}

func (this *TypeLib) GetTypeInfoCount() int {
	count := this.p.GetTypeInfoCount()
	return int(count)
}

func (this *TypeLib) GetTypeInfo(index int) *TypeInfo {
	var pti *win32.ITypeInfo
	hr := this.p.GetTypeInfo(uint32(index), &pti)
	win32.ASSERT_SUCCEEDED(hr)
	return NewTypeInfo(pti)
}
