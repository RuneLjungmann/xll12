// args.h - Arguments to register an Excel add-in
// Copyright (c) KALX, LLC. All rights reserved. No warranty made.
#pragma once
#include <algorithm>
#include <string>
#include "defines.h"
#include "excel.h"

namespace xll {

	/// See https://msdn.microsoft.com/en-us/library/office/bb687900.aspx
	enum ARG {
		ModuleText,   // from xlGetName
		Procedure,    // C function
		TypeText,     // return type and arg codes 
		FunctionText, // Excel function
		ArgumentText, // Ctrl-Shift-A text
		MacroType,    // function, 2 for macro, 0 for hidden
		Category,     // for function wizard
		ShortcutText, // single character for Ctrl-Shift-char shortcut
		HelpTopic,    // filepath!HelpContextID or http://host/path!0
		FunctionHelp, // for function wizard
		ArgumentHelp, // 1-based index
	};

	using xcstr = const wchar_t*;

	/// <summary>Prepare an array suitible for <c>xlfRegister</c></summary>
	class Args {
		mutable OPER12 args;
        OPER12 ArgumentName_;
        OPER12 ArgumentDefault_;
        xcstr documentation = nullptr;
        xcstr remarks = nullptr;
        xcstr examples = nullptr;
	public:
		/// Name of Excel add-in
		static OPER12& XlGetName()
		{
			static OPER12 hModule;

			if (hModule.type() != xltypeStr) {
				hModule = Excel(xlGetName);
			}

			return hModule;
		}
		
		/// <summary>Number of function arguments</summary>
		int Arity() const
		{
			return ArgumentName_.size();
		}
		OPER12 RegisterId() const
		{
			//return Excel(xlfEvaluate, Excel(xlfConcatenate, OPER12(L"="), args[ARG::FunctionText]));
            return Excel(xlfEvaluate, args[ARG::FunctionText]);
        }
		/// For use as Excelv(xlfRegister, Args(....))
		operator const OPER12&() const
		{
			return args;
		}

		/// Common default.
		Args()
			: args(1, ARG::ArgumentHelp)
		{
			std::fill(args.begin(), args.end(), OPER12(xltype::Nil));
		}
		Args(const Args&) = default;
		Args& operator=(const Args&) = default;
		~Args()
		{ }

		/// Macro
		Args(xcstr Procedure, xcstr FunctionText)
			: Args()
		{
			//args[ARG::ModuleText] = XlGetName();
			args[ARG::Procedure] = Procedure;
			args[ARG::FunctionText] = FunctionText;
			args[ARG::MacroType] = OPER12(2);
		}
		/// Function
		Args(xcstr TypeText, xcstr Procedure, xcstr FunctionText)
			: Args()
		{
			//args[ARG::ModuleText] = XlGetName();
			args[ARG::Procedure] = Procedure;
			args[ARG::TypeText] = TypeText;
			args[ARG::FunctionText] = FunctionText;
			args[ARG::MacroType] = OPER12(1);
		}
        /// Documentation
        Args(xcstr _documentation)
            : Args()
        {
            // needed for Key()
            args[ARG::FunctionText] = L"*";
            documentation = _documentation;
            args[ARG::MacroType] = OPER(-1);
        }

		Args& ModuleText(const OPER& moduleText)
		{
			args[ARG::ModuleText] = moduleText;

			return *this;
		}

		/// Set the name of the C/C++ function to be called.
		Args& Procedure(xcstr procedure)
		{
			args[ARG::Procedure] = procedure;

			return *this;
		}
        const OPER& Procedure() const
        {
            return args[ARG::Procedure];
        }

		/// Specify the return type and argument types of the function.
		Args& TypeText(xcstr typeText)
		{
			args[ARG::TypeText] = typeText;

			return *this;
		}
		/// Specify the name of the function or macro to be used by Excel.
		Args& FunctionText(xcstr functionText)
		{
			args[ARG::FunctionText] = functionText;

			return *this;
		}
		const OPER& FunctionText() const
		{
			return args[ARG::FunctionText];
		}
		/// Specify the macro type of the function.
		/// Use 1 for functions, 2 for macros, and 0 for hidden functions. 
		Args& MacroType(int macroType)
		{
			args[ARG::MacroType] = macroType;

			return *this;
		}
        bool isFunction() const
        {
            return args[ARG::MacroType] == 1;
        }
        bool isMacro() const
        {
            return args[ARG::MacroType] == 2;
        }
        bool isHidden() const
        {
            return args[ARG::MacroType] == 0;
        }
        bool isDocumentation() const
        {
            return args[ARG::MacroType] == -1; // special type
        }

        /// Hide the name of the function from Excel.
		Args& Hidden()
		{
			return MacroType(0);
		}
		/// Set the category to be used in the function wizard.
		Args& Category(xcstr category)
		{
			args[ARG::Category] = category;

			return *this;
		}
		const OPER& Category() const
		{
			return args[ARG::Category];
		}
		/// Specify the shortcut text for calling the function.
		Args& ShortcutText(wchar_t shortcutText)
		{
			args[ARG::ShortcutText] = OPER(&shortcutText, 1);

			return *this;
		}
		/// Specify the help topic to be used in the Function Wizard.
		/// This must have the format...
		Args& HelpTopic(xcstr helpTopic)
		{
			args[ARG::HelpTopic] = helpTopic;

			return *this;
		}
		/// Specify the function help displayed in the Functinon Wizard.
		Args& FunctionHelp(xcstr functionHelp)
		{
			args[ARG::FunctionHelp] = functionHelp;

			return *this;
		}
        const OPER& FunctionHelp() const
        {
            return args[ARG::FunctionHelp];
        }
        /// Specify individual argument help in the Function Wizard.
		Args& ArgumentHelp(COL i, xcstr argumentHelp)
		{
			ensure (i != 0);

			if (args.size() < ARG::ArgumentHelp + i)
				args.resize(1, ARG::ArgumentHelp + i);

			auto n = ARG::ArgumentHelp + i - 1;
			args[n] = argumentHelp;

			return *this;
		}
        const OPER& ArgumentHelp(int i) const
        {
            return args[ARG::ArgumentHelp + i - 1];
        }
        OPER ArgumentName(int i) const
        {
			ensure(i <= Arity());

            return ArgumentName_[i - 1];
        }
		OPER ArgumentDefault(int i) const
		{
			ensure(i <= Arity());

			return ArgumentDefault_[i - 1];
		}
		/// Add an individual argument.
		Args& Arg(xcstr type, xcstr name, xcstr helpText = nullptr, xcstr Default = nullptr)
		{
			OPER& Type = args[ARG::TypeText];
			Type &= type;
			
			OPER& Text = args[ARG::ArgumentText];
			if (Text.isStr())
				Text &= L", ";
			Text &= name;

            ArgumentName_.push_back(OPER(name));
			
            RW n = Arity();
			if (helpText && *helpText) {
				ArgumentHelp(n, helpText);
            }
            if (Default && *Default) {
                ArgumentDefault_.resize(n,1);
                ArgumentDefault_[n - 1] = Default;
            }

			return *this;
		}
		/// Argument modifiers
		Args& Threadsafe()
		{
			args[ARG::TypeText] &= XLL_THREAD_SAFE;

			return *this;
		}
		int isThreadsafe()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_THREAD_SAFE)).isNum();
		}
		Args& Uncalced()
		{
			args[ARG::TypeText] &= XLL_UNCALCED;

			return *this;
		}
		int isUncalced()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_UNCALCED)).isNum();
		}
		Args& Volatile()
		{
			args[ARG::TypeText] &= XLL_VOLATILE;

			return *this;
		}
		int isVolatile()
		{
			return Excel(xlfFind, args[ARG::TypeText], OPER(XLL_VOLATILE)).isNum();
		}

		/// Convenience function for number types.
		Args& Num(xcstr name, xcstr helpText = nullptr)
		{
			return Arg(XLL_DOUBLE, name, helpText);
		}
		// Str ...

		Args& Documentation(xcstr _documentation)
		{
            documentation = _documentation;

			return *this;
		}
        xcstr Documentation() const
        {
            return documentation;
        }
        Args& Remarks(xcstr _remarks)
        {
            remarks = _remarks;

            return *this;
        }
        xcstr Remarks() const
        {
            return remarks;
        }
        Args& Examples(xcstr _examples)
        {
            examples = _examples;

            return *this;
        }
        xcstr Examples() const
        {
            return examples;
        }
        OPER Syntax() const
        {
            return FunctionText() & OPER(L"(") & args[ARG::ArgumentText] & OPER(L")");
        }

        OPER Key() const
        {
			//??? Add prefix???
			return FunctionText();
        }

	// Simple hash function
        static OPER hash_string(const wchar_t* s, unsigned n)
        {
            static int A = 54059; /* a prime */
            static int B = 76963; /* another prime */
            static int C = 86969; /* yet another prime */
            static int FIRST = 37; /* also prime */
            int h = FIRST;
            while (n--) {
                h = (h * A) ^ (s[0] * B);
                s++;
            }

            return Excel(xlfText, OPER(abs(h)), OPER(L"General")); // or return h % C;
        }

        // Integer hash used in help files.
        OPER TopicId() const
        {
            auto key = Key();

            return hash_string(key.val.str + 1, key.val.str[0]);
        }

        OPER Guid() const 
        {
            return Excel(xlfDec2hex, TopicId()) & OPER(L"-0000-0000-0000-000000000000");
        }

		/// Register an add-in function or macro
		OPER Register() const
		{
            if (isDocumentation()) {
                return OPER(1); // Do not register if documentation only.
            }

            OPER name = XlGetName();
            args[ARG::ModuleText] = name;
            
            if (documentation && *documentation) {
                OPER chm = Excel(xlfSubstitute, name, OPER(L".xll"), OPER(L".chm!"));
                //chm &= OPER(L"0");
                chm &= TopicId();
                args[ARG::HelpTopic] = chm;
            }

			OPER oResult = Excelv(xlfRegister, args);
			if (oResult.isErr()) {
				OPER oError(L"Failed to register: ");
				oError &= args[ARG::FunctionText];
				oError &= L"/";
				oError &= args[ARG::Procedure];
                oError &= L"\nDid you forget to #pragma XLLEXPORT a function?";
				
                Excel(xlcAlert, oError);
			}

			return oResult;
		}
        /// Unregister and add-in function or macro
        int Unregister() const
        {
            if (isDocumentation())
                return TRUE;

            return Excel(xlfUnregister, RegisterId()) == true;
        }
	};
	// semantic alias
	using Function = Args;
	using Macro = Args;
    using Documentation = Args;
	// backwards compatibility
	using ArgsX = Args;
	using FunctionX = Args;
	using MacroX = Args;
	using Args12 = Args;
	using Function12 = Args;
	using Macro12 = Args;

	/// Array appropriate for xlfRegister.
	/// Use like <c>Excelv(xlfRegister, Arguments(...))</c>
	inline OPER Arguments(
		xcstr Procedure,        // C function
		xcstr TypeText,         // return type and arg codes 
		xcstr FunctionText,     // Excel function
		xcstr ArgumentText = 0, // Ctrl-Shift-A text
		int   MacroType = 1,    // function, 2 for macro, 0 for hidden
		xcstr Category = 0,     // for function wizard
		xcstr ShortcutText = 0, // single character for Ctrl-Shift-char shortcut
		xcstr HelpTopic = 0,    // filepath!HelpContextID or http://host/path!0
		xcstr FunctionHelp = 0, // for function wizard
		xcstr ArgumentHelp1 = 0,
		xcstr ArgumentHelp2 = 0,
		xcstr ArgumentHelp3 = 0,
		xcstr ArgumentHelp4 = 0,
		xcstr ArgumentHelp5 = 0,
		xcstr ArgumentHelp6 = 0,
		xcstr ArgumentHelp7 = 0,
		xcstr ArgumentHelp8 = 0,
		xcstr ArgumentHelp9 = 0
	)
	{
		OPER args(Args::XlGetName());
		args.push_back(OPER(Procedure));
		args.push_back(OPER(TypeText));
		args.push_back(OPER(FunctionText));
		args.push_back(OPER(ArgumentText));
		args.push_back(OPER(MacroType));
		args.push_back(OPER(Category));
		args.push_back(OPER(ShortcutText));
		args.push_back(OPER(HelpTopic));
		args.push_back(OPER(FunctionHelp));
		args.push_back(OPER(ArgumentHelp1));
		args.push_back(OPER(ArgumentHelp2));
		args.push_back(OPER(ArgumentHelp3));
		args.push_back(OPER(ArgumentHelp4));
		args.push_back(OPER(ArgumentHelp5));
		args.push_back(OPER(ArgumentHelp6));
		args.push_back(OPER(ArgumentHelp7));
		args.push_back(OPER(ArgumentHelp8));
		args.push_back(OPER(ArgumentHelp9));

		return args;
	}
	/*
	template<class R, class... Args>
	Args AutoRegister(R(*)(Args...), xcstr Procedure, xcstr FunctionText)
	{

	}
	*/

	/*
	// Convert __FUNCDNAME__ to arguments for xlfRegister
	inline Args Demangle(const wchar_t* F)
	{
		static std::map<wchar_t,const OPER12> arg_map = {
			{ L'F', OPER12(XLL_SHORT) },
			{ L'G', OPER12(XLL_WORD) }, // also USHORT
			{ L'H', OPER12(XLL_BOOL) },
			{ L'J', OPER12(XLL_LONG) },
			{ L'N', OPER12(XLL_DOUBLE) },
		};

		// C to Excel naming convention
		auto function_text = [](OPER12 o) {
			ensure (o.type() == xltypeStr);
			for (int i = 1; i <= o.val.str[0]; ++i) {
				if (o.val.str[i] == L'_')
					o.val.str[i] = L'.';
				else
					o.val.str[i] = ::towupper(o.val.str[i]);
			}

			return o;
		};

		Args args;
		args.MacroType(1); // function
								  
		// "?foo@@YGNN@Z"

		ensure (F && *F == '?');
		auto E = wcschr(F, L'@');
		ensure (E);
		args.Procedure(OPER12(F, E - F));
		args.FunctionText(function_text(OPER12(F + 1, E - F - 1)));

		F = E;
		ensure (*++F == L'@');
		ensure (*++F == L'Y');
		ensure (*++F == L'G');
		ensure (arg_map.find(*++F) != arg_map.end());
		args.TypeText(arg_map[*F]);
		while (*++F != L'@') {
			// if (*F == L'P') { pointer...
			ensure (arg_map.find(*F) != arg_map.end());
//			type = Excel(xlfConcatenate, type, OPER12(arg_map[*F]));
		}

		return args;
	}
	*/
} // xll
