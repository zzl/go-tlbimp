package codegen

import (
	"fmt"
	"github.com/zzl/go-tlbimp/typelib"
	"github.com/zzl/go-tlbimp/utils"
	"github.com/zzl/go-win32api/win32"
	"io/ioutil"
	"os"
	"path"
	"strconv"
	"strings"
)

type Generator struct {
	TypeLib    *typelib.TypeLib
	RefLibMap  map[string]*typelib.TypeLib
	OutputPath string

	codeMap map[string]string

	ownClassSet    map[string]bool
	sourceClassSet map[string]bool

	refClassMap     map[string]string //name:pkg
	usedRefClassMap map[string]string
}

func (this *Generator) Generate() {

	this.prepareRefInfo()
	this.prepareOwnInfo()

	this.OutputPath = strings.ReplaceAll(this.OutputPath, "\\", "/")
	this.cleanOutputDir()

	this.codeMap = make(map[string]string)
	tiCount := this.TypeLib.GetTypeInfoCount()
	for n := 0; n < tiCount; n++ {
		ti := this.TypeLib.GetTypeInfo(n)
		this.genType(ti)
	}

	this.writeCodes()
}

func isWin32Type(typeName string) bool {
	switch typeName {
	case "IUnknown", "ISequentialStream", "IStream", "IDispatch",
		"DISPPARAMS", "EXCEPINFO":
		return true
	}
	return false
}

func (this *Generator) prepareOwnInfo() {
	this.ownClassSet = make(map[string]bool)
	this.sourceClassSet = make(map[string]bool)
	tiCount := this.TypeLib.GetTypeInfoCount()
	for n := 0; n < tiCount; n++ {
		ti := this.TypeLib.GetTypeInfo(n)
		if ti.Kind == win32.TKIND_COCLASS ||
			ti.Kind == win32.TKIND_INTERFACE ||
			ti.Kind == win32.TKIND_DISPATCH {
			if !isWin32Type(ti.Name) {
				this.ownClassSet[utils.CapName(ti.Name)] = true
			}
		}
		if ti.Kind == win32.TKIND_COCLASS {
			for _, it := range ti.ImplTypes {
				if it.Source {
					this.sourceClassSet[utils.CapName(it.Name)] = true
				}
			}
		}
	}
}

func (this *Generator) prepareRefInfo() {
	this.refClassMap = make(map[string]string)
	this.usedRefClassMap = make(map[string]string)

	for pkg, tlb := range this.RefLibMap {
		tiCount := tlb.GetTypeInfoCount()
		for n := 0; n < tiCount; n++ {
			ti := tlb.GetTypeInfo(n)
			if ti.Kind == win32.TKIND_COCLASS ||
				ti.Kind == win32.TKIND_INTERFACE ||
				ti.Kind == win32.TKIND_DISPATCH {
				name := utils.CapName(ti.Name)
				if !this.ownClassSet[name] {
					this.refClassMap[name] = pkg
				}
			}
		}
	}
}

func (this *Generator) writeCodes() {
	pkgName := path.Base(this.OutputPath)
	for name, code := range this.codeMap {
		code := "package " + pkgName + "\n\n" + genImports(code) + code
		filePath := path.Join(this.OutputPath, name+".go")
		ioutil.WriteFile(filePath, []byte(code), os.ModePerm)
	}
	refsCode := this.genRefsCode()
	if refsCode != "" {
		filePath := path.Join(this.OutputPath, "refs.go")
		ioutil.WriteFile(filePath, []byte(refsCode), os.ModePerm)
	}
}

func (this *Generator) genRefsCode() string {
	if len(this.usedRefClassMap) == 0 {
		return ""
	}
	pkgSet := make(map[string]bool)     //pkg
	classMap := make(map[string]string) //class:pkg_last
	for className, pkg := range this.usedRefClassMap {
		pkgSet[pkg] = true
		pos := strings.LastIndexByte(pkg, '/')
		pkgLast := pkg
		if pos != -1 {
			pkgLast = pkg[pos+1:]
		}
		classMap[className] = pkgLast
	}
	var code string
	pkgName := path.Base(this.OutputPath)
	code += "package " + pkgName + "\n\n"
	code += "import (\n"
	for pkg, _ := range pkgSet {
		code += "\t\"" + pkg + "\"\n"
	}
	code += ")\n\n"

	for className, pkgLast := range classMap {
		code += "type " + className + " = " + pkgLast + "." + className + "\n"
		if !isWin32Type(className) {
			code += "var New" + className + " = " + pkgLast + ".New" + className + "\n\n"
		}
	}
	return code
}

func (this *Generator) cleanOutputDir() {
	fis, _ := ioutil.ReadDir(this.OutputPath)
	for _, fi := range fis {
		c := fi.Name()[0]
		if c < '0' || c > '9' {
			filePath := path.Join(this.OutputPath, fi.Name())
			os.Remove(filePath)
		}
	}
}

func genImports(code string) string {
	var imports string
	if strings.Contains(code, "win32") {
		imports += "\t\"github.com/zzl/go-win32api/win32\"\n"
	}
	if strings.Contains(code, "com.") {
		imports += "\t\"github.com/zzl/go-com/com\"\n"
	}
	if strings.Contains(code, "ole.") {
		imports += "\t\"github.com/zzl/go-com/ole\"\n"
	}
	if strings.Contains(code, "syscall") {
		imports += "\t\"syscall\"\n"
	}
	if strings.Contains(code, "unsafe") {
		imports += "\t\"unsafe\"\n"
	}
	if strings.Contains(code, "time.Time") {
		imports += "\t\"time\"\n"
	}
	if strings.Contains(code, "runtime.") {
		imports += "\t\"runtime\"\n"
	}
	if strings.Contains(code, "reflect.") {
		imports += "\t\"reflect\"\n"
	}
	if imports != "" {
		imports = "import (\n" + imports + ")\n\n"
	}
	return imports
}

func (this *Generator) genType(ti *typelib.TypeInfo) {
	switch ti.Kind {
	case win32.TKIND_ENUM:
		this.genEnum(ti)
	case win32.TKIND_RECORD:
		this.genStruct(ti)
	case win32.TKIND_UNION:
		this.genUnion(ti)
	case win32.TKIND_ALIAS:
		this.genAlias(ti)
	case win32.TKIND_INTERFACE:
		if strings.HasSuffix(ti.Name, "Handler") { //?
			this.genHandlerInterface(ti)
		} else {
			this.genInterface(ti)
		}
	case win32.TKIND_DISPATCH:
		if this.sourceClassSet[utils.CapName(ti.Name)] {
			this.genSourceDispInterface(ti)
		} else {
			this.genDispInterface(ti)
		}
	case win32.TKIND_COCLASS:
		this.genCoClass(ti)
	}
}

func (this *Generator) genAlias(ti *typelib.TypeInfo) {
	if ti.Name == "GUID" {
		return
	}

	code := this.codeMap["types"]
	code += "// alias " + ti.Name + "\n"
	name := utils.CapName(ti.Name)
	code += "type " + name + " = " + ti.RelType.Name + "\n\n"
	this.codeMap["types"] = code
}

//
type simpleFieldInfo struct {
	name string
	typ  string
}

func (this *Generator) genUnion(ti *typelib.TypeInfo) {

	if strings.Contains(ti.Name, "MIDL_") && strings.Contains(ti.Name, "WinType") {
		return
	}

	size, alignSize := ti.Size, ti.Align
	var dataFields []simpleFieldInfo

	embedFieldIndex := -1
	for n, f := range ti.Fields {
		if f.Name == "Anonymous" {
			fSize, fAlign := f.Type.Size, f.Type.Align
			if fSize == size {
				embedFieldIndex = n
			} else {
				_ = fAlign
				//?
			}
			break
		}
	}
	if embedFieldIndex != -1 {
		f := ti.Fields[embedFieldIndex]
		dataFields = append(dataFields, simpleFieldInfo{
			name: "",
			typ:  f.Type.Name,
		})
	} else {
		var elemType string
		switch alignSize {
		case 1:
			elemType = "byte"
		case 2:
			elemType = "uint16"
		case 4:
			elemType = "uint32"
		case 8:
			elemType = "uint64"
		default:
			panic("?")
		}
		elemCount := size / alignSize
		typ := fmt.Sprintf("[%d]%s", elemCount, elemType)
		dataFields = append(dataFields, simpleFieldInfo{
			name: "Data",
			typ:  typ,
		})
	}

	var unionFields []simpleFieldInfo
	for n, f := range ti.Fields {
		if n == embedFieldIndex {
			continue
		}
		unionFields = append(unionFields, simpleFieldInfo{
			name: utils.CapName(f.Name),
			typ:  f.Type.Name,
		})
	}

	//
	code := this.codeMap["types"]
	code += "// union " + ti.Name + "\n"
	name := utils.CapName(ti.Name)
	code += "type " + name + " struct {\n"
	for _, f := range dataFields {
		code += "\t" + f.name + " " + f.typ + "\n"
	}
	code += "}\n\n"

	for _, f := range unionFields {
		code += "func (this *" + name + ") " + f.name + "() *" + f.typ + " {\n"
		code += "\treturn (*" + f.typ + ")(unsafe.Pointer(this))\n"
		code += "}\n\n"
		code += "func (this *" + name + ") " + f.name + "Val () " + f.typ + " {\n"
		code += "\treturn *(*" + f.typ + ")(unsafe.Pointer(this))\n"
		code += "}\n\n"
	}
	this.codeMap["types"] = code
}

func (this *Generator) genStruct(ti *typelib.TypeInfo) {
	code := this.codeMap["types"]
	code += "// struct " + ti.Name + "\n"
	name := utils.CapName(ti.Name)
	code += "type " + name + " struct {\n"
	count := ti.FieldCount
	for n := 0; n < count; n++ {
		f := ti.GetField(n)
		code += "\t" + utils.CapName(f.Name) + " " + f.Type.Name + "\n"
	}
	code += "}\n\n"
	this.codeMap["types"] = code
}

func (this *Generator) genEnum(ti *typelib.TypeInfo) {
	code := this.codeMap["enums"]

	code += "// enum " + ti.Name + "\n"
	code += "var " + utils.CapName(ti.Name) + " = struct {\n"
	count := ti.FieldCount
	for n := 0; n < count; n++ {
		f := ti.GetField(n)
		code += "\t" + utils.CapName(f.Name) + " " + f.Type.Name + "\n"
	}
	code += "}{\n"
	for n := 0; n < count; n++ {
		f := ti.GetField(n)
		sValue := fmt.Sprintf("%v", f.Value)
		code += "\t" + utils.CapName(f.Name) + ": " + sValue + ",\n"
	}
	code += "}\n\n"

	this.codeMap["enums"] = code
}

func (this *Generator) genDispInterface(ti *typelib.TypeInfo) {
	className := utils.CapName(ti.Name)
	code := this.codeMap[className]

	sIid, _ := win32.GuidToStr(&ti.Guid)
	iidExpr := utils.BuildGuidExpr(sIid)
	code += "// " + sIid + "\n"
	code += "var IID_" + className + " = " + iidExpr + "\n\n"

	code += "type " + className + " struct {\n"
	code += "\tole.OleClient\n"
	code += "}\n\n"

	code += "func New" + className + "(pDisp *win32.IDispatch, addRef bool, scoped bool) *" + className + " {\n"
	code += "\tp := &" + className + "{ole.OleClient{pDisp}}\n"
	code += "\tif addRef {\n"
	code += "\t\tpDisp.AddRef()\n"
	code += "\t}\n"
	code += "\tif scoped {\n"
	code += "\t\tcom.AddToScope(p)\n"
	code += "\t}\n"
	code += "\treturn p\n"
	code += "}\n\n"

	//
	code += "func " + className + "FromVar(v ole.Variant) *" + className + " {\n"
	code += "\treturn New" + className + "(v.PdispValVal(), false, false)\n"
	code += "}\n\n"

	//
	code += "func (this *" + className + ") IID() *syscall.GUID {\n"
	code += "\treturn &IID_" + className + "\n"
	code += "}\n\n"

	//
	code += "func (this *" + className + ") GetIDispatch(addRef bool) *win32.IDispatch {\n"
	code += "\tif addRef {\n"
	code += "\t\tthis.AddRef()\n"
	code += "\t}\n"
	code += "\treturn this.IDispatch\n"
	code += "}\n\n"

	//
	var fromFuncIndex int
	if ti.DualInterface != nil {
		fromFuncIndex = 7
	}
	//
	count := ti.FuncCount

	setMethods := make(map[string]bool)
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		if f.Flags.PropPut || f.Flags.PropPutRef {
			setMethods["Set"+utils.CapName(f.Name)] = true
		}
	}
	var colItemType, itemReturnType *typelib.VarType
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		if f.Id == int32(win32.DISPID_VALUE) {
			colItemType = f.ReturnType
			break
		}
		if f.Name == "Item" || f.Name == "GetItem" {
			itemReturnType = f.ReturnType
		}
	}
	if colItemType == nil {
		colItemType = itemReturnType
	}
	if colItemType == nil {
		colItemType = &typelib.VarType{Name: "*IUnknown", Pointer: true}
	}

	setNames := make(map[string]bool)
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		var methodType string
		if f.Flags.PropGet {
			methodType = "PropGet"
		} else if f.Flags.PropPut {
			if setNames[f.Name] {
				continue
			}
			methodType = "PropPut"
			setNames[f.Name] = true
		} else if f.Flags.PropPutRef {
			if setNames[f.Name] {
				continue
			}
			methodType = "PropPutRef"
			setNames[f.Name] = true
		} else {
			methodType = "Call"
		}
		code += this.genDispMethod(f, className, methodType, setMethods)

		if f.Id == win32.DISPID_NEWENUM {
			code += this.genForEachEnum(f, className, colItemType)
		}
	}
	this.codeMap[className] = code
}

func (this *Generator) genSourceDispInterface(ti *typelib.TypeInfo) {
	interfaceName := utils.CapName(ti.Name)
	code := this.codeMap[interfaceName]

	sIid, _ := win32.GuidToStr(&ti.Guid)
	iidExpr := utils.BuildGuidExpr(sIid)
	code += "// " + sIid + "\n"
	code += "var IID_" + interfaceName + " = " + iidExpr + "\n\n"

	code += "type " + interfaceName + "DispInterface interface {\n"

	var fromFuncIndex int
	if ti.DualInterface != nil {
		fromFuncIndex = 7
	}
	//
	count := ti.FuncCount

	superFuncs := collectInheritedFuncs(ti.Super)
	superMethods := make(map[string]bool)
	for _, f := range superFuncs {
		superMethods[utils.CapName(f.Name)] = true
	}

	setMethods := make(map[string]bool)

	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		if f.Flags.PropPut || f.Flags.PropPutRef {
			setMethods["Set"+utils.CapName(f.Name)] = true
		}
	}
	fNames := make([]string, count)
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		fName := utils.CapName(f.Name)
		if f.Flags.PropPut || f.Flags.PropPutRef {
			fName = "Set" + fName
		} else if setMethods[fName] || superMethods[fName] {
			fName += "_"
		}
		fNames[n] = fName
	}

	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		sDispId := this.genDispId(f)
		_ = sDispId
		code += "\t" + this.genSourceDispMethod(f, fNames[n])
	}

	code += "}\n\n"

	//
	code += "type " + interfaceName + "Handlers struct {\n"
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		code += "\t" + fNames[n] + " " + this.genSourceDispMethod(f, "func")
	}
	code += "}\n\n"

	//
	dispImplClass := interfaceName + "DispImpl"
	code += "type " + dispImplClass + " struct {\n"
	code += "\tHandlers " + interfaceName + "Handlers\n"
	code += "}\n\n"

	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		code += "" + this.genSourceDispMethodImpl(f, dispImplClass, fNames[n])
	}

	//
	implClass := interfaceName + "Impl"
	code += "type " + implClass + " struct {\n"
	code += "\tole.IDispatchImpl\n" //?
	code += "\tDispImpl " + interfaceName + "DispInterface\n"
	code += "}\n\n"

	code += "func (this *" + implClass + ") QueryInterface(" +
		"riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {\n"
	code += "\tif *riid == IID_" + interfaceName + " {\n"
	code += "\t\tthis.AssignPpvObject(ppvObject)\n"
	code += "\t\tthis.AddRef()\n"
	code += "\t\treturn win32.S_OK\n"
	code += "\t}\n"
	code += "\treturn this.IDispatchImpl.QueryInterface(riid, ppvObject)\n"
	code += "}\n\n"

	code += "func (this *" + implClass + ") Invoke(" +
		"dispIdMember int32, riid *syscall.GUID, lcid uint32,\n"
	code += "\twFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,\n"
	code += "\tpExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {\n"

	code += "\tvar unwrapActions ole.Actions\n"
	code += "\tdefer unwrapActions.Execute()\n"
	code += "\tswitch dispIdMember {\n"
	for n := fromFuncIndex; n < count; n++ {
		f := ti.GetFunc(n)
		code += "\tcase " + fmt.Sprintf("%v", f.Id) + ":\n"
		if len(f.Params) > 0 {
			code += "\t\tvArgs, _ := ole.ProcessInvokeArgs(pDispParams, " +
				strconv.Itoa(len(f.Params)) + ")\n"
		}
		condition := false
		if f.Flags.PropPut || f.Flags.PropPutRef {
			code += "\t\tif wFlags == win32.DISPATCH_PROPERTYPUT || " +
				"wFlags == win32.DISPATCH_PROPERTYPUTREF {\n"
			condition = true
		} else if f.Flags.PropGet {
			code += "\t\tif wFlags == win32.DISPATCH_PROPERTYGET{\n"
			condition = true
		}
		if condition {
			code += "\t"
		}
		code += this.genDispFuncInvoke(f, fNames[n]) + "\n"
		if condition {
			code += "\t\t}\n"
		}
	}
	code += "\t}\n"
	code += "\treturn win32.E_NOTIMPL\n"
	code += "}\n\n"

	code += "type " + interfaceName + "ComObj struct {\n"
	code += "\tole.IDispatchComObj\n"
	code += "}\n\n"

	code += "func New" + interfaceName + "ComObj(" +
		"dispImpl " + interfaceName + "DispInterface, scoped bool) *" + interfaceName + "ComObj {\n"

	code += "\tcomObj := com.NewComObj[" + interfaceName + "ComObj](\n" +
		"\t\t&" + interfaceName + "Impl {DispImpl: dispImpl})\n"

	code += "\tif scoped {\n"
	code += "\t\tcom.AddToScope(comObj)\n"
	code += "\t}\n"
	code += "\treturn comObj\n"
	code += "}\n\n"

	//
	this.codeMap[interfaceName] = code
}

func (this *Generator) genDispFuncInvoke(f *typelib.FuncInfo, fName string) string {
	var code string
	goReturnType := this.mapOleTypeToGoType(f.ReturnType, true)
	pCount := len(f.Params)
	_ = pCount
	sArgs := ""
	for n, p := range f.Params {
		_ = p
		if n > 0 {
			sArgs += ", "
		}
		goParamType := this.mapOleTypeToGoType(p.Type, false)
		vArg := "vArgs[" + strconv.Itoa(n) + "]"

		aName := "p" + strconv.Itoa(n+1)
		sArgs += aName
		if p.Type.Pointer && goParamType != "string" {
			code += "\t\t" + aName + " := " +
				"(" + goParamType + ")(" + vArg + ".ToPointer())\n"
		} else if goParamType == "ole.Variant" {
			code += "\t\t" + aName + ", _ := " + vArg + "\n"
		} else {
			code += "\t\t" + aName + ", _ := " +
				"" + vArg + ".To" + utils.CapName(goParamType) + "()\n"
		}
	}
	if goReturnType != "" {
		code += "\t\tret := "
	} else {
		code += "\t\t"
	}
	code += "this.DispImpl." + fName + "("
	code += sArgs
	code += ")\n"
	if goReturnType != "" {
		code += "\t\tole.SetVariantParam((*ole.Variant)(pVarResult), ret, &unwrapActions)\n"
	}
	code += "\t\treturn win32.S_OK"
	return code
}

func (this *Generator) genSourceDispMethod(f *typelib.FuncInfo,
	fName string) string {

	var code string
	goReturnType := this.mapOleTypeToGoType(f.ReturnType, true)
	code += fName + "("

	for n, p := range f.Params {
		if n > 0 {
			code += ", "
		}
		pName := utils.UncapName(p.Name)
		pName = utils.SafeGoName(pName)
		code += pName + " "
		goParamType := this.mapOleTypeToGoType(p.Type, false)
		code += goParamType
	}
	code += ") " + goReturnType + "\n"
	return code
}

func (this *Generator) genSourceDispMethodImpl(f *typelib.FuncInfo,
	implClass string, fName string) string {

	var code string
	goReturnType := this.mapOleTypeToGoType(f.ReturnType, true)
	code += "func (this *" + implClass + ") " + fName + "("

	for n, p := range f.Params {
		if n > 0 {
			code += ", "
		}
		pName := utils.UncapName(p.Name)
		pName = utils.SafeGoName(pName)
		code += pName + " "
		goParamType := this.mapOleTypeToGoType(p.Type, false)
		code += goParamType
	}
	if goReturnType != "" {
		code += ") " + goReturnType + " {\n"
	} else {
		code += ") {\n"
	}

	//
	code += "\tif this.Handlers." + fName + " != nil {\n"
	code += "\t\t"
	if goReturnType != "" {
		code += "return "
	}
	code += "this.Handlers." + fName + "("
	for n, p := range f.Params {
		if n > 0 {
			code += ", "
		}
		pName := utils.UncapName(p.Name)
		pName = utils.SafeGoName(pName)
		code += pName
	}
	code += ")\n"
	code += "\t}\n"
	if goReturnType != "" {
		if goReturnType == "win32.HRESULT" {
			code += "\treturn win32.E_NOTIMPL\n"
		} else {
			code += "\tvar ret " + goReturnType + "\n"
			code += "\treturn ret\n"
		}
	}

	code += "}\n\n"
	return code
}

func (this *Generator) genForEachEnum(f *typelib.FuncInfo,
	className string, itemType *typelib.VarType) string {
	var code string
	itemTypeName := this.mapOleTypeToGoType(itemType, true)
	code += "func (this *" + className + ") " +
		"ForEach(action func(item " + itemTypeName + ") bool) {\n"

	code += "\tpEnum := this." + utils.CapName(f.Name) + "()\n"
	code += "\tvar pEnumVar *win32.IEnumVARIANT\n"
	code += "\tpEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))\n"
	code += "\tdefer pEnumVar.Release();\n"
	code += "\tfor {\n"

	code += "\t\tvar c uint32\n"
	code += "\t\tvar v ole.Variant\n"
	code += "\t\tpEnumVar.Next(1, (*win32.VARIANT)(&v), &c)\n"
	code += "\t\tif c == 0 {\n"
	code += "\t\t\tbreak\n"
	code += "\t\t}\n"

	code += "\t\t"
	if itemType.Pointer && itemTypeName != "string" {
		code += "pItem := (" + itemTypeName + ")(v.ToPointer())\n"
	} else if itemTypeName == "ole.Variant" {
		code += "pItem := v\n"
	} else {
		code += "pItem, _ := v.To" + utils.CapName(itemTypeName) + "()\n"
	}

	code += "\t\tret := action(pItem)\n"
	if itemType.Pointer {
		code += "\t\tv.Clear()\n"
	}
	code += "\t\tif !ret {\n"
	code += "\t\t\tbreak\n"
	code += "\t\t}\n"
	code += "\t}\n"
	code += "}\n\n"

	return code
}

func (this *Generator) genDispMethod(f *typelib.FuncInfo, className string,
	methodType string, setMethods map[string]bool) string {

	var code string
	sDispId := this.genDispId(f)
	goReturnType := this.mapOleTypeToGoType(f.ReturnType, true)

	fName := utils.CapName(f.Name)

	switch methodType {
	case "PropGet":
		//
	case "PropPut":
		fName = "Set" + fName
	case "PropPutRef":
		fName = "Set" + fName
	case "Call":
		if setMethods[fName] {
			fName += "_"
		}
	}
	if fName == "QueryInterface" {
		fName += "_"
	}

	optParamCount := 0
	var optArgsVarName string
	for n, p := range f.Params {
		if !p.Flags.Optional {
			continue
		}
		if optParamCount == 0 {
			optArgsVarName = className + "_" + fName + "_OptArgs"
			code += "var " + optArgsVarName + "= []string{\n\t"
		} else if (optParamCount)%4 == 0 && n != len(f.Params)-1 {
			code += "\n\t"
		}
		code += "\"" + p.Name + "\", "
		optParamCount += 1
	}
	if optParamCount > 0 {
		code += "\n}\n\n"
	}

	code += "func (this *" + className + ") " + fName + "("

	if className == "Range" && fName == "SetItem" {
		println("?")
	}

	var reqParamNames []string
	for _, p := range f.Params {
		if p.Flags.Optional {
			break
		}
		if reqParamNames != nil {
			code += ", "
		}
		pName := utils.UncapName(p.Name)
		pName = utils.SafeGoName(pName)
		reqParamNames = append(reqParamNames, pName)
		code += pName + " "
		goParamType := this.mapOleTypeToGoType(p.Type, false)
		code += goParamType
	}

	if optParamCount > 0 {
		if reqParamNames != nil {
			code += ", "
		}
		code += "optArgs ...interface{}"
	}

	code += ") " + goReturnType + " {\n"

	if optParamCount > 0 {
		code += "\toptArgs = ole.ProcessOptArgs(" + optArgsVarName + ", optArgs)\n"
	}

	code += "\tretVal := this." + methodType + "(" + sDispId
	if reqParamNames != nil {
		code += ", []interface{}{" + strings.Join(reqParamNames, ", ") + "}"
	} else {
		code += ", nil"
	}
	if optParamCount > 0 {
		code += ", optArgs..."
	}
	code += ")\n"

	//
	if goReturnType == "ole.Variant" {
		code += "\tcom.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))\n"
	}

	//
	code += "\t" + this.genDispReturnCode(f.ReturnType, goReturnType) + "\n"
	code += "}\n\n"
	return code
}

func (this *Generator) genDispId(f *typelib.FuncInfo) string {
	var sDispId string
	if f.Id < 0 {
		sDispId = fmt.Sprintf("%v", f.Id)
	} else {
		sDispId = fmt.Sprintf("0x%08x", uint32(f.Id))
	}
	return sDispId
}

func (this *Generator) mapOleTypeToGoType(varType *typelib.VarType, forReturn bool) string {
	oleType := varType.Name
	if oleType == "" { //void
		return "" //?
	}
	var goType string
	if oleType == "win32.VARIANT_BOOL" {
		goType = "bool"
	} else if oleType == "win32.BSTR" {
		goType = "string"
	} else if oleType == "ole.Date" {
		goType = "time.Time"
	} else if oleType == "win32.HRESULT" {
		goType = "com.Error"
	} else if oleType == "*win32.IUnknown" {
		goType = "*com.UnknownClass"
	} else if oleType == "*win32.IDispatch" {
		goType = "*ole.DispatchClass"
	} else if oleType == "win32.PWSTR" {
		goType = "string"
	} else if oleType[0] == '*' && isWin32Type(oleType[1:]) {
		goType = "*win32." + oleType[1:]
	} else if varType.Pointer && varType.RefType.Interface &&
		!this.ownClassSet[oleType[1:]] {
		if this.refClassMap[oleType[1:]] != "" {
			this.usedRefClassMap[oleType[1:]] = this.refClassMap[oleType[1:]]
			goType = oleType //
		} else {
			if varType.RefType.DispInterface {
				goType = "*ole.DispatchClass"
			} else {
				goType = "*com.UnknownClass"
			}
		}
	} else if varType.Pointer && varType.RefType.Pointer &&
		varType.RefType.RefType != nil && varType.RefType.RefType.Interface &&
		!this.ownClassSet[oleType[2:]] {
		if this.refClassMap[oleType[2:]] != "" {
			this.usedRefClassMap[oleType[2:]] = this.refClassMap[oleType[2:]]
			goType = oleType //
		} else {
			if varType.RefType.DispInterface {
				goType = "**ole.DispatchClass"
			} else {
				goType = "**com.UnknownClass"
			}
		}
	} else {
		goType = oleType
	}
	if goType == "win32.VARIANT" {
		if forReturn {
			return "ole.Variant"
		} else {
			return "interface{}"
		}
	} else if goType == "*win32.VARIANT" {
		return "*ole.Variant"
	}
	return goType
}

func (this *Generator) genDispReturnCode(varType *typelib.VarType, goType string) string {
	oleType := varType.Name
	if oleType == "" {
		return "_= retVal"
	}
	if goType[0] == '*' {
		if varType.RefType != nil && varType.RefType.Interface {
			if goType == "*com.UnknownClass" {
				return "return com.NewUnknownClass(retVal.PunkValVal(), true)"
			} else if goType == "*ole.DispatchClass" {
				return "return ole.NewDispatchClass(retVal.PdispValVal(), true)"
			} else if varType.RefType.DispInterface {
				return "return New" + goType[1:] + "(retVal.PdispValVal(), false, true)"
			} else {
				return "return New" + goType[1:] + "(retVal.PunkValVal(), false, true)"
			}
		}
	}

	castExpr := strings.Replace(varType.PVarCastExpr, "$", "retVal", 1)
	switch oleType {
	case "ole.Date":
		return "return " + castExpr + ".ToGoTime()"
	case "win32.BSTR":
		return "return win32.BstrToStrAndFree(" + castExpr + ")"
	case "win32.HRESULT":
		return "return com.NewError(" + castExpr + ")"
	case "win32.VARIANT_BOOL":
		return "return " + castExpr + " != win32.VARIANT_FALSE"
	}
	return "return " + castExpr
}

func (this *Generator) genCoClass(ti *typelib.TypeInfo) {
	className := utils.CapName(ti.Name)
	code := this.codeMap[className]

	var implTi *typelib.ImplType
	var sourceTi *typelib.ImplType
	for _, it := range ti.ImplTypes {
		if it.Default {
			if it.Source {
				sourceTi = it
			} else {
				implTi = it
			}
		}
	}
	implClass := utils.CapName(implTi.Name)

	sIid, _ := win32.GuidToStr(&ti.Guid)
	iidExpr := utils.BuildGuidExpr(sIid)
	code += "var CLSID_" + className + " = " + iidExpr + "\n\n"

	code += "type " + className + " struct {\n"
	code += "\t" + implClass + "\n"
	code += "}\n\n"

	if implTi.DispInterface {
		code += "func New" + className + "(pDisp *win32.IDispatch, addRef bool, scoped bool) *" + className + " {\n"
		code += "\tp := &" + className + "{" + implClass + "{ole.OleClient{pDisp}}}\n"

		code += "\tif addRef {\n"
		code += "\t\tpDisp.AddRef()\n"
		code += "\t}\n"

		code += "\tif scoped {\n"
		if implClass == "*win32.IUnknown" {
			code += "\t\tcom.AddToScope(p)\n"
		} else {
			code += "\t\tcom.AddToScope(p)\n"
		}
		code += "\t}\n"
		code += "\treturn p\n"
		code += "}\n\n"

		//
		code += "func New" + className + "FromVar(v ole.Variant, addRef bool, scoped bool) *" + className + " {\n"
		code += "\treturn New" + className + "(v.PdispValVal(), addRef, scoped)\n"
		code += "}\n\n"
	} else {
		code += "func New" + className + "(pUnk *win32.IUnknown, addRef bool, scoped bool) *" + className + " {\n"
		code += "\tp := (*" + className + ")(unsafe.Pointer(pUnk))\n"
		code += "\tif addRef {\n"
		code += "\t\tpUnk.AddRef()\n"
		code += "\t}\n"
		code += "\tif scoped {\n"
		if implClass == "*win32.IUnknown" {
			code += "\t\tcom.AddToScope(p)\n"
		} else {
			code += "\t\tcom.AddToScope(p)\n"
		}
		code += "\t}\n"
		code += "\treturn p\n"
		code += "}\n\n"
	}

	//
	code += "func New" + className + "Instance(scoped bool) (*" + className + ", error) {\n"
	code += "\tvar p *"
	if implTi.DispInterface {
		code += "win32.IDispatch\n"
	} else {
		code += "win32.IUnknown\n"
	}
	//
	code += "\thr := win32.CoCreateInstance(&CLSID_" + className + ", nil, \n" +
		"\t\twin32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,\n" +
		"\t\t&IID_" + implClass + ", unsafe.Pointer(&p))\n"
	code += "\tif win32.FAILED(hr) {\n"
	code += "\t\treturn nil, com.NewError(hr)\n"
	code += "\t}\n"
	code += "\treturn New" + className + "(p, false, scoped), nil\n"
	code += "}\n\n"

	//
	if sourceTi != nil {
		sourceClass := utils.CapName(sourceTi.Name)
		code += "func (this *" + className + ") " +
			"RegisterEventHandlers(handlers " + sourceClass + "Handlers) uint32 {\n"
		code += "\tvar cpc *win32.IConnectionPointContainer\n"
		code += "\thr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\tvar cp *win32.IConnectionPoint\n"
		code += "\thr = cpc.FindConnectionPoint(&IID_" + sourceClass + ", &cp)\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\tdispImpl := &" + sourceClass + "DispImpl{Handlers: handlers}\n"
		code += "\tdisp := New" + sourceClass + "ComObj(dispImpl, false)\n"
		code += "\t\n"
		code += "\tvar cookie uint32\n"
		code += "\thr = cp.Advise(disp.IUnknown(), &cookie)\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\tdisp.Release()\n"
		code += "\tcp.Release()\n"
		code += "\tcpc.Release()\n"
		code += "\treturn cookie\n"
		code += "}\n\n"

		code += "func (this *" + className + ") " +
			"UnRegisterEventHandlers(cookie uint32) {\n"
		code += "\tvar cpc *win32.IConnectionPointContainer\n"
		code += "\thr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\tvar cp *win32.IConnectionPoint\n"
		code += "\thr = cpc.FindConnectionPoint(&IID_" + sourceClass + ", &cp)\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\thr = cp.Unadvise(cookie)\n"
		code += "\twin32.ASSERT_SUCCEEDED(hr)\n"
		code += "\n"
		code += "\tcp.Release()\n"
		code += "\tcpc.Release()\n"
		code += "}\n\n"
	}
	this.codeMap[className] = code
}

func (this *Generator) genInterface(ti *typelib.TypeInfo) {
	className := utils.CapName(ti.Name)
	if isWin32Type(className) {
		return
	}
	code := this.codeMap[className]

	sIid, _ := win32.GuidToStr(&ti.Guid)
	iidExpr := utils.BuildGuidExpr(sIid)
	code += "// " + sIid + "\n"
	code += "var IID_" + className + " = " + iidExpr + "\n\n"

	//
	superClassName := utils.CapName(ti.Super.Name)
	if isWin32Type(superClassName) {
		superClassName = "win32." + superClassName
	}

	//
	code += "type " + className + " struct {\n"

	code += "\t" + superClassName + "\n"
	code += "}\n\n"

	code += "func New" + className + "(pUnk *win32.IUnknown, addRef bool, scoped bool) *" + className + " {\n"
	code += "\tp := (*" + className + ")(unsafe.Pointer(pUnk))\n"

	code += "\tif addRef {\n"
	code += "\t\tpUnk.AddRef()\n"
	code += "\t}\n"

	code += "\tif scoped {\n"
	code += "\t\tcom.AddToScope(p)\n"
	code += "\t}\n"
	code += "\treturn p\n"
	code += "}\n\n"

	//
	code += "func (this *" + className + ") IID() *syscall.GUID {\n"
	code += "\treturn &IID_" + className + "\n"
	code += "}\n\n"

	//
	fCount := ti.FuncCount
	funcs := ti.Funcs

	superFuncCount := len(collectInheritedFuncs(ti.Super))

	setMethods := make(map[string]bool)
	for n := 0; n < fCount; n++ {
		f := funcs[n]
		if f.Flags.PropPut || f.Flags.PropPutRef {
			setMethods["Set"+utils.CapName(f.Name)] = true
		}
	}
	for n := 0; n < fCount; n++ {
		f := funcs[n]
		fIndex := n + superFuncCount
		if f.Flags.PropGet {
			code += this.genPropGet(className, fIndex, f)
		} else if f.Flags.PropPut {
			code += this.genPropPut(className, fIndex, f, setMethods)
		} else if f.Flags.PropPutRef {
			code += this.genPropPutRef(className, fIndex, f)
		} else {
			code += this.genMethod(className, fIndex, f, setMethods)
		}
	}
	this.codeMap[className] = code
}

//
func (this *Generator) genHandlerInterface(ti *typelib.TypeInfo) {
	class := utils.CapName(ti.Name)
	superClass := utils.CapName(ti.Super.Name)

	code := this.codeMap[class]

	sIid, _ := win32.GuidToStr(&ti.Guid)
	iidExpr := utils.BuildGuidExpr(sIid)
	code += "// " + sIid + "\n"
	code += "var IID_" + class + " = " + iidExpr + "\n\n"

	code += "type " + class + " struct {\n"
	code += "\t"
	if isWin32Type(superClass) {
		code += "win32."
	}
	code += superClass + "\n"
	code += "}\n\n"

	code += "type " + class + "Interface interface {\n"
	code += "\t"
	if isWin32Type(superClass) {
		code += "win32."
	}
	code += superClass + "Interface" + "\n"
	fCount := ti.FuncCount

	fNames := make([]string, fCount)
	for n := 0; n < fCount; n++ {
		f := ti.GetFunc(n)
		fName := utils.CapName(f.Name)
		fNames[n] = fName
	}

	for n := 0; n < fCount; n++ {
		f := ti.GetFunc(n)
		code += "\t" + this.genFunc("", fNames[n], n, f.Params, f.ReturnType, true) + "\n"
	}

	code += "}\n\n"

	classImpl := class + "Impl"
	code += "type " + classImpl + " struct {\n"
	code += "\t"
	if isWin32Type(superClass) {
		code += "com."
	}
	code += superClass + "Impl\n" //?
	code += "\tRealObject " + class + "Interface\n"
	code += "}\n\n"

	code += "func (this *" + classImpl + ") SetRealObject(obj interface{}) {\n"
	code += "\tthis.RealObject = obj.(" + class + "Interface)\n"
	code += "}\n\n"

	code += "func (this *" + classImpl + ") QueryInterface(" +
		"riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {\n"
	code += "\tif *riid == IID_" + class + " {\n"
	code += "\t\tthis.AssignPpvObject(ppvObject)\n"
	code += "\t\tthis.AddRef()\n"
	code += "\t\treturn win32.S_OK\n"
	code += "\t}\n"
	code += "\treturn this." + superClass + "Impl.QueryInterface(riid, ppvObject)\n"
	code += "}\n\n"

	for n := 0; n < fCount; n++ {
		f := ti.GetFunc(n)
		code += "func (this *" + classImpl + ") " + fNames[n] + "("
		for m, p := range f.Params {
			if m > 0 {
				code += ", "
			}
			pName := utils.UncapName(p.Name)
			pName = utils.SafeGoName(pName)
			goParamType := this.mapOleTypeToGoType(p.Type, false)
			code += pName + " " + goParamType
		}
		code += ") "
		goReturnType := this.mapOleTypeToGoType(f.ReturnType, true)
		if goReturnType != "" {
			code += goReturnType + " "
		}
		code += "{\n"
		if goReturnType != "" {
			code += "\tvar ret " + goReturnType + "\n"
			code += "\treturn ret\n"
		}
		code += "}\n"
	}

	//
	code += "type " + class + "Vtbl struct {\n"
	code += "\t"
	if isWin32Type(superClass) {
		code += "win32."
	}
	code += superClass + "Vtbl\n"
	for n := 0; n < ti.FuncCount; n++ {
		f := ti.Funcs[n]
		code += "\t" + utils.CapName(f.Name) + " uintptr\n"
	}
	code += "}\n\n"

	//
	code += "type " + class + "ComObj struct {\n"
	code += "\t"
	if isWin32Type(superClass) {
		code += "com."
	}
	code += superClass + "ComObj\n"
	code += "}\n\n"

	code += "func (this *" + class + "ComObj) impl() " + class + "Interface {\n"
	code += "\treturn this.Impl().(" + class + "Interface)\n"
	code += "}\n\n"

	for n := 0; n < fCount; n++ {
		f := ti.Funcs[n]
		fName := utils.CapName(f.Name)

		//
		code += "func (this *" + class + "ComObj) "
		code += fName + "("

		var pNames []string
		for n, p := range f.Params {
			if n > 0 {
				code += ", "
			}
			pName := utils.UncapName(p.Name)
			pName = utils.SafeGoName(pName)
			pNames = append(pNames, pName)
			code += pName + " "

			goParamType := this.mapOleTypeToGoType(p.Type, false)
			if goParamType == "string" {
				goParamType = p.Type.Name
			}
			code += goParamType
		}

		code += ") uintptr {\n"

		//
		code += "\treturn (uintptr)(this.impl()." + fName + "("
		for m, p := range f.Params {
			pName := pNames[m]
			if m > 0 {
				code += ", "
			}
			goParamType := this.mapOleTypeToGoType(p.Type, false)
			if goParamType == "string" {
				if p.Type.Name == "win32.BSTR" {
					code += "win32.BstrToStr(" + pName + ")"
				} else if p.Type.Name == "win32.PWSTR" {
					code += "win32.PwstrToStr(" + pName + ")"
				} else {
					panic("?")
				}
			} else {
				code += pName
			}
		}
		code += "))\n"
		code += "}\n\n"
	}

	//
	code += "var _p" + class + "Vtbl *" + class + "Vtbl\n\n"
	code += "func (this *" + class +
		"ComObj) BuildVtbl(lock bool) *" + class + "Vtbl {\n"
	code += "\tif lock {\n"
	code += "\t\tcom.MuVtbl.Lock()\n"
	code += "\t\tdefer com.MuVtbl.Unlock()\n"
	code += "}\n"

	code += "\tif _p" + class + "Vtbl != nil {\n"
	code += "\t\treturn _p" + class + "Vtbl\n"
	code += "\t}\n"
	code += "\t_p" + class + "Vtbl = &" + class + "Vtbl{\n"
	code += "\t\t" + superClass + "Vtbl: *this." +
		superClass + "ComObj.BuildVtbl(false),\n"
	for n := 0; n < ti.FuncCount; n++ {
		f := ti.Funcs[n]
		fName := utils.CapName(f.Name)
		code += "\t\t" + fName + ":\tsyscall.NewCallback((*" +
			class + "ComObj)." + fName + "),\n"
	}
	code += "\t}\n"
	code += "\treturn _p" + class + "Vtbl\n"
	code += "}\n\n"

	code += "func (this *" + class + "ComObj) " +
		class + "() *" + class + "{\n"

	code += "\treturn (*" + class + ")(unsafe.Pointer(this))\n"
	code += "}\n\n"

	//
	code += "func (this *" + class + "ComObj) GetVtbl() *win32.IUnknownVtbl {\n"
	code += "\treturn &this.BuildVtbl(true).IUnknownVtbl\n"
	code += "}\n\n"

	//
	code += "func New" + class + "ComObj(impl " +
		class + "Interface, scoped bool) *" + class + "ComObj {\n"

	code += "\tcomObj := com.NewComObj[" + class + "ComObj](impl)\n"

	code += "\tif scoped {\n"
	code += "\t\tcom.AddToScope(comObj)\n"
	code += "\t}\n"
	code += "\treturn comObj\n"
	code += "}\n\n"

	code += "func New" + class + "(impl " + class + "Interface) *" + class + " {\n"
	code += "\treturn New" + class + "ComObj(impl, true)." + class + "()"
	code += "}\n\n"

	if fCount == 1 && fNames[0] == "Invoke" {
		code += "//\n"
		code += "type " + class + "ByFuncImpl struct {\n"
		code += "\t" + classImpl + "\n"

		f := ti.GetFunc(0)
		code += "\thandlerFunc func " + this.genFunc("", "", 0, f.Params, f.ReturnType, true) + "\n"
		code += "}\n"

		//
		code += this.genFunc(class+"ByFuncImpl", "Invoke",
			0, f.Params, f.ReturnType, true) + "{\n"
		code += "\treturn this.handlerFunc("
		for m, p := range f.Params {
			if m > 0 {
				code += ", "
			}
			pName := utils.UncapName(p.Name)
			pName = utils.SafeGoName(pName)
			code += pName
		}
		code += ")\n"
		code += "}\n\n"

		//
		code += "func New" + class + "ByFunc(handlerFunc func " +
			this.genFunc("", "", 0, f.Params, f.ReturnType, true) +
			", scoped bool) *" + class + " {\n"
		code += "\timpl := &" + class + "ByFuncImpl{handlerFunc: handlerFunc}\n"
		code += "\treturn New" + class + "ComObj(impl, scoped)." + class + "()" + "\n"
		code += "}\n\n"
	}

	this.codeMap[class] = code
}

//
func collectInheritedFuncs(superTi *typelib.TypeInfo) []*typelib.FuncInfo {
	var funcs []*typelib.FuncInfo
	if superTi.Super != nil {
		funcs = append(funcs, collectInheritedFuncs(superTi.Super)...)
	}
	funcs = append(funcs, superTi.Funcs...)
	return funcs
}

func (this *Generator) genPropGet(className string, fIndex int, f *typelib.FuncInfo) string {
	fName := "Get" + utils.CapName(f.Name)
	code := this.genFunc(className, fName, fIndex, f.Params, f.ReturnType, false)
	return code
}

func (this *Generator) genPropPut(className string, fIndex int, f *typelib.FuncInfo,
	setMethods map[string]bool) string {

	fName := "Set" + utils.CapName(f.Name)
	setMethods[fName] = true
	code := this.genFunc(className, fName, fIndex, f.Params, f.ReturnType, false)
	return code
}

func (this *Generator) genPropPutRef(className string, fIndex int, f *typelib.FuncInfo) string {
	fName := "Set" + utils.CapName(f.Name)
	code := this.genFunc(className, fName, fIndex, f.Params, f.ReturnType, false)
	return code
}

func (this *Generator) genMethod(className string, fIndex int, f *typelib.FuncInfo,
	setMethods map[string]bool) string {

	fName := utils.CapName(f.Name)
	if setMethods[fName] {
		fName += "_"
	}

	code := this.genFunc(className, fName, fIndex, f.Params, f.ReturnType, false)
	return code
}

func (this *Generator) genFunc(className string, fName string, fIndex int,
	params []*typelib.ParamInfo, returnType *typelib.VarType, noBody bool) string {

	var code string
	goReturnType := this.mapOleTypeToGoType(returnType, true)

	if className != "" {
		code += "func (this *" + className + ") "
	} else {
		code += ""
	}
	code += fName + "("

	var pNames []string
	var pTypes []string
	for n, p := range params {
		if n > 0 {
			code += ", "
		}
		pName := utils.UncapName(p.Name)
		pName = utils.SafeGoName(pName)
		pNames = append(pNames, pName)
		code += pName + " "

		goParamType := this.mapOleTypeToGoType(p.Type, false)
		pTypes = append(pTypes, goParamType)
		code += goParamType
	}

	code += ") " + goReturnType

	if noBody {
		//code += "\n"
		return code
	}

	code += " {\n"

	code += "\taddr := (*this.LpVtbl)[" + strconv.Itoa(fIndex) + "]\n"
	code += "\t"
	if goReturnType != "" {
		code += "ret, _, _ :="
	} else {
		code += "_, _, _ ="
	}
	code += " syscall.SyscallN(addr, uintptr(unsafe.Pointer(this))"
	var outInterfaceParams []string
	for n, pName := range pNames {
		code += ", "
		param := params[n]
		pType := pTypes[n]
		if pType == "bool" {
			if param.Type.Name == "VARIANT_BOOL" {
				code += "uintptr(^(*(*VARIANT_BOOL)" +
					"(unsafe.Pointer(&" + pName + ")) - 1))"
			} else {
				code += "uintptr(*(*uint8)(unsafe.Pointer(&" + pName + ")))"
			}
		} else if pType[0] == '*' {
			code += "uintptr(unsafe.Pointer(" + pName + "))"
			if pType[1] == '*' && params[n].Type.RefType.RefType != nil &&
				(params[n].Type.RefType.RefType.Interface ||
					params[n].Type.RefType.RefType.DispInterface) {
				outInterfaceParams = append(outInterfaceParams, pName)
			}
		} else if pType == "uintptr" {
			code += pName
		} else if pType == "string" {
			code += "uintptr(win32.StrToPointer(" + pName + "))"
		} else if param.Type.Struct {
			if param.Type.Size > utils.PtrSize {
				code += "(uintptr)(unsafe.Pointer(&" + pName + "))"
			} else {
				code += "*(*uintptr)(unsafe.Pointer(&" + pName + "))"
			}
		} else {
			code += "uintptr(" + pName + ")"
		}
	}
	code += ")\n"
	if outInterfaceParams != nil {
		code += "\tif com.CurrentScope != nil {\n"
	}
	for _, outParam := range outInterfaceParams {
		code += "\t\tcom.CurrentScope.Add(unsafe.Pointer(&(*" + outParam + ").IUnknown))\n"
	}
	if outInterfaceParams != nil {
		code += "\t}\n"
	}
	if goReturnType != "" {
		code += "\t" + this.genReturnCode(returnType, goReturnType) + "\n"
	}
	code += "}\n\n"

	return code
}

func (this *Generator) genReturnCode(typ *typelib.VarType, goType string) string {
	var castExpr string
	switch goType {
	case "bool":
		castExpr = "ret != 0"
	case "string":
		castExpr = "win32.BstrToStrAndFree(win32.BSTR(unsafe.pointer(ret))"
	case "time.Time":
		castExpr = "ole.Date(ret).ToGoTime()"
	case "com.Error":
		castExpr = "com.Error(ret)"
	case "uintptr":
		castExpr = "ret"
	case "unsafe.Pointer":
		castExpr = "unsafe.Pointer(ret)"
	}
	if castExpr != "" {
		return "return " + castExpr
	}
	if typ.Struct {
		return "return *(*" + goType + ")(unsafe.Pointer(ret))"
	}
	if goType[0] == '*' {
		return "return (" + goType + ")(unsafe.Pointer(ret))"
	}
	return "return " + goType + "(ret)"
}
