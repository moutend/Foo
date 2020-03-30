package main

import (
	"fmt"
	"log"
	"os"

	"github.com/moutend/Foo/pkg/com"

	"github.com/go-ole/go-ole"
)

func main() {
	log.SetPrefix("error: ")
	log.SetFlags(0)

	if err := run(os.Args); err != nil {
		log.Fatal(err)
	}
}

func run(args []string) error {
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
		return err
	}
	defer ole.CoUninitialize()

	var foo *com.IFoo

	if err := com.CoCreateInstance(com.CLSID_Foo, 0, com.CLSCTX_ALL, com.IID_IFoo, &foo); err != nil {
		return err
	}

	fmt.Println("@@@instance created")

	defer foo.Stop()
	defer foo.Release()

	err := foo.Start()
	fmt.Println("Called IFoo::Start", err)

	return nil
}
