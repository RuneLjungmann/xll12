// oper.h - C++ wrapper for OPER
#pragma once
#define NOMINMAX

#include <Windows.h>
#include "XLCALL.H"
#include <algorithm>
#include <limits>
#include <memory>
#include <string>

namespace xll {

	struct OPER12 : public XLOPER12 
	{
		friend void swap(OPER12& a, OPER12& b)
		{
			using std::swap;

			swap(a.xltype, b.xltype);
			swap(a.val, b.val);
		}

		OPER12()
		{
			xltype = xltypeNil;
		}
		OPER12(const OPER12& o)
		{
			xltype = o.xltype;
			if (xltype == xltypeStr) {
				allocate_str(o.val.str[0]);
				copy_str(o.val.str + 1);
			}
			else if (xltype == xltypeMulti) {
				allocate_multi(o.val.array.rows, o.val.array.columns);
				std::copy(o.begin(), o.end(), begin());
			}
			else if (xltype == xltypeBigData) {
				ensure(!"not implemented");
			}
			else {
				val = o.val;
			}
		}
		OPER12(OPER12&& o)
		{
			xltype = o.xltype;
			val = o.val;

			o.xltype = xltypeNil;
		}
		OPER12& operator=(OPER12 o)
		{
			swap(*this, o);

			return *this;
		}
		~OPER12()
		{
			if (xltype == xltypeStr) {
				deallocate_str();
			}
			else if (xltype == xltypeMulti) {
				destroy_multi();
				deallocate_multi();
			}
		}

		bool operator==(const OPER12& o) const
		{
			if (xltype != o.xltype)
				return false;

			switch (xltype&~(xlbitXLFree|xlbitDLLFree)) {
			case xltypeNum:
				return val.num == o.val.num;
			case xltypeStr:
				return val.str[0] == o.val.str[0] && 0 == wcsncmp(val.str + 1, o.val.str + 1, val.str[0]);
			case xltypeBool:
				return val.xbool == o.val.xbool;
			case xltypeRef:
				return false;
			case xltypeErr:
				return val.err == o.val.err;
			case xltypeFlow:
				return false; // ancient macro programming artifact
			case xltypeMulti:
				return val.array.rows == o.val.array.rows
					&& val.array.columns == o.val.array.columns
					&& true; //std::equal(begin(), end(), o.begin(), o.end());
			case xltypeMissing:
			case xltypeNil:
				return true;
			case xltypeInt:
				return val.w == o.val.w;
			case xltypeBigData:
				return false; //? check val.bigdata.lpbData
			}

			return false;
		}
		bool operator!=(const OPER12& o) const
		{
			return !operator==(o);
		}
		// bool operator<(...)
		

		RW rows() const
		{
			return xltype == xltypeMulti ? val.array.rows 
				: xltype == xltypeNil ? 0 : 1;
		}
		COL columns() const
		{
			return xltype == xltypeMulti ? val.array.columns 
				: xltype == xltypeNil ? 0 : 1;
		}
		size_t size() const
		{
			return rows() * columns();
		}

		OPER12* begin()
		{
			return xltype == xltypeMulti ? static_cast<OPER12*>(val.array.lparray) : this;
		}
		const OPER12* begin() const
		{
			return xltype == xltypeMulti ? static_cast<const OPER12*>(val.array.lparray) : this;
		}
		OPER12* end()
		{
			return xltype == xltypeMulti ? static_cast<OPER12*>(val.array.lparray + size()) : this + 1;
		}
		const OPER12* end() const
		{
			return xltype == xltypeMulti ? static_cast<const OPER12*>(val.array.lparray + size()) : this + 1;
		}

		// Num
		explicit OPER12(const double& num)
		{
			xltype = xltypeNum;
			val.num = num;
		}
		OPER12& operator=(const double& num)
		{
			return *this = OPER12(num);
		}
		bool operator==(const double& num) const
		{
			return xltype == xltypeNum && val.num == num;
		}
		operator double()
		{
			ensure (xltype == xltypeNum || xltype == xltypeInt);

			return xltype == xltypeNum ? val.num : val.w;
		}

		// Str
		explicit OPER12(const XCHAR* str)
			: OPER12(str, wcslen(str))
		{ }
		OPER12(const XCHAR* str, size_t len)
		{
			xltype = xltypeStr;
			allocate_str(wcslen(str));
			copy_str(str);
		}
		OPER12& operator=(const std::wstring& str)
		{
			return *this = OPER12(str.c_str(), str.length());
		}
		bool operator==(const XCHAR* str) const
		{
			return 0 == wcsncmp(str, val.str + 1, val.str[0]);
		}

		// Multi
		OPER12(RW rw, COL col)
		{
			xltype = xltypeMulti;
			val.array.lparray = nullptr;//alloc_oper.allocate(rw*col);
		}
		OPER12& operator[](size_t i)
		{
			if (xltype != xltypeMulti) {
				ensure (i == 0);
				return *this;
			}

			return static_cast<OPER12&>(val.array.lparray[i]);
		}
		const OPER12& operator[](size_t i) const
		{
			if (xltype != xltypeMulti) {
				ensure (i == 0);
				return *this;
			}

			return static_cast<const OPER12&>(val.array.lparray[i]);
		}
		OPER12& operator()(RW i, COL j)
		{
			return operator[](i + j*columns());
		}
		const OPER12& operator()(RW i, COL j) const
		{
			return operator[](i + j*columns());
		}

		// Int
		OPER12(const int& w)
		{
			xltype = xltypeInt;
			val.w = w;
		}
		bool operator==(const int& w) const
		{
			return xltype == xltypeInt && val.w == w
				|| xltype == xltypeNum && val.num == w;
		}
	private:
		void allocate_str(size_t len)
		{
			ensure (len < std::numeric_limits<XCHAR>::max());
			val.str = static_cast<XCHAR*>(::malloc((1 + len)*sizeof(XCHAR)));
			ensure (val.str);
			val.str[0] = static_cast<XCHAR>(len);
			xltype = xltypeStr;
		}
		void copy_str(const XCHAR* str)
		{
			wcsncpy(val.str + 1, str, val.str[0]);
		}
		void deallocate_str()
		{
			ensure (xltype == xltypeStr);

			::free(val.str);
		}
		void allocate_multi(RW rw, COL col)
		{
			// check rw, col size?
			val.array.lparray = static_cast<XLOPER12*>(::malloc(rw*col*sizeof(XLOPER12)));
			ensure (val.array.lparray);
			val.array.rows = rw;
			val.array.columns = col;
			xltype = xltypeMulti;
		}
		void uninitialied_copy_multi(const OPER12* i)
		{
			ensure (xltype == xltypeMulti);

			for (auto& o : *this) {
				new (static_cast<void*>(&o)) OPER12(*i++);
			}
		}
		void destroy_multi()
		{
			ensure (xltype == xltypeMulti);

			for (auto& o : *this)
				o.~OPER12();
		}
		void deallocate_multi()
		{
			ensure (xltype == xltypeMulti);

			::free(val.array.lparray);
		}
	};

} // xll
