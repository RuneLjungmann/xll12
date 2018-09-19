#pragma once
#include "../xll/xll.h"

namespace xll {

    __declspec(align(8)) class base {
        int data_;
    public:
        base(int data = 0)
            : data_(data)
        { }
        virtual ~base() // for RTTI
        { }
        int value(void) const
        {
            return data_;
        }
    };

    __declspec(align(8)) class derived : public base {
        int data_;
    public:
        derived(int bdata = 0, int ddata = 0)
            : base(bdata), data_(ddata)
        { }
        ~derived()
        { }
        int value2(void) const
        {
            return data_;
        }
    };

}