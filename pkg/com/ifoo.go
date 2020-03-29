package com

import (
	"unsafe"

	"github.com/go-ole/go-ole"
)

type IFoo struct {
	ole.IDispatch
}

type IFooVtbl struct {
	ole.IDispatchVtbl
	Start       uintptr
	Stop        uintptr
	DoSomething uintptr
}

func (v *IFoo) VTable() *IFooVtbl {
	return (*IFooVtbl)(unsafe.Pointer(v.RawVTable))
}

func (v *IFoo) Start() error {
	return fStart(v)
}

func (v *IFoo) Stop() error {
	return fStop(v)
}

func (v *IFoo) DoSomething() error {
	return fDoSomething(v)
}
