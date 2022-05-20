package main

import (
	"flag"
	"fmt"
	"github.com/zzl/go-tlbimp/codegen"
	"github.com/zzl/go-tlbimp/typelib"
	"github.com/zzl/go-tlbimp/utils"
	"os"
	"strings"
)

var tlbPath string
var outputDir string
var sRefTlbs string
var sRefPkgs string

func main() {

	flag.StringVar(&tlbPath, "tlb", "", "target tlb file path")
	flag.StringVar(&outputDir, "out-dir", "", "output directory")

	flag.StringVar(&sRefTlbs, "imp-tlbs", "", "import tlb file paths(; separated)")
	flag.StringVar(&sRefPkgs, "imp-pkgs", "", "import package names(; separated)")

	flag.Parse()
	if tlbPath == "" || outputDir == "" {
		flag.Usage()
		return
	}

	if !utils.FileExists(tlbPath) {
		println("Target tlb not found: " + tlbPath)
		return
	}

	tlb, err := typelib.NewTypeLibFromFile(tlbPath)
	if err != nil {
		println("Failed to load " + tlbPath)
		return
	}

	if !utils.DirExists(outputDir) {
		println("Output dir does not exist: " + outputDir)
		print("Create now? (Y/N) ")
		var answer string
		fmt.Scanf("%s", &answer)
		if answer != "Y" && answer != "y" {
			return
		} else {
			err = os.MkdirAll(outputDir, 0700)
			if err != nil {
				println("Failed to create output dir.")
				return
			}
		}
	}

	refTlbPaths := strings.Split(strings.TrimRight(sRefTlbs, ";"), ";")
	if len(refTlbPaths) == 1 && refTlbPaths[0] == "" {
		refTlbPaths = nil
	}
	refPkgs := strings.Split(strings.TrimRight(sRefPkgs, ";"), ";")
	if len(refPkgs) == 1 && refPkgs[0] == "" {
		refPkgs = nil
	}
	if len(refTlbPaths) != len(refPkgs) {
		println("Number of imp-tlbs and imp-pkgs do not match.")
		return
	}

	var generator codegen.Generator
	generator.TypeLib = tlb
	generator.OutputPath = outputDir

	generator.RefLibMap = make(map[string]*typelib.TypeLib)
	for n, refTlbPath := range refTlbPaths {
		if !utils.FileExists(refTlbPath) {
			println("Ref tlb not found: " + refTlbPath)
			return
		}
		refTlb, err := typelib.NewTypeLibFromFile(refTlbPath)
		if err != nil {
			println("Failed to load " + refTlbPath)
			return
		}
		refPkg := refPkgs[n]
		generator.RefLibMap[refPkg] = refTlb
	}

	generator.Generate()
	println("Done.")
}
