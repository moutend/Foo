// +build !windows

package com

import (
	"github.com/go-ole/go-ole"
)

func fStart(v *IFoo) error {
	return ole.NewError(ole.E_NOTIMPL)
}

func fStop(v *IFoo) error {
	return ole.NewError(ole.E_NOTIMPL)
}

func fDoSomething(v *IFoo) error {
	return ole.NewError(ole.E_NOTIMPL)
}
