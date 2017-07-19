#using <mscorlib.dll>
#include <vcclr.h>
#include <memory>
#include <powershell_cpp/string.hpp>


#pragma managed
using namespace System;
using namespace System::Management::Automation;


ref class PSOutHelper
{
public:
    void OnDataAddedDebug(Object ^sender, DataAddedEventArgs ^args)
    {
        ConsoleColor original = Console::ForegroundColor;
        Console::ForegroundColor = ConsoleColor::Blue;
        Console::WriteLine(args);
        Console::ForegroundColor = original;
    }

    void OnDataAddedProgress(Object ^sender, DataAddedEventArgs ^args)
    {
        ConsoleColor original = Console::ForegroundColor;
        Console::ForegroundColor = ConsoleColor::Green;
        Console::WriteLine(args);
        Console::ForegroundColor = original;
    }

    void OnDataAddedVerbose(Object ^sender, DataAddedEventArgs ^args)
    {
        Console::WriteLine(args);
    }

    void OnDataAddedWarning(Object ^sender, DataAddedEventArgs ^args)
    {
        ConsoleColor original = Console::ForegroundColor;
        Console::ForegroundColor = ConsoleColor::Yellow;
        Console::WriteLine(args);
        Console::ForegroundColor = original;
    }

    void OnDataAddedError(Object ^sender, DataAddedEventArgs ^args)
    {
        ConsoleColor original = Console::ForegroundColor;
        Console::ForegroundColor = ConsoleColor::Red;
        Console::WriteLine(args);
        Console::ForegroundColor = original;
    }
};

namespace ps {
    /// PowerShell interpreter
    class Interpreter
    {
    public:

        Interpreter()
        {
            h = gcnew PSOutHelper();
        }

        /// Execute <paramref="script_code" /> synchronously
        void exec_script_sync(const string &script_code)
        {
            PowerShell ^ps_interpreter = PowerShell::Create();
            ps_interpreter->Streams->Debug->DataAdded += gcnew EventHandler<DataAddedEventArgs ^>(h, &PSOutHelper::OnDataAddedDebug);
            ps_interpreter->Streams->Progress->DataAdded += gcnew EventHandler<DataAddedEventArgs ^>(h, &PSOutHelper::OnDataAddedProgress);
            ps_interpreter->Streams->Verbose->DataAdded += gcnew EventHandler<DataAddedEventArgs ^>(h, &PSOutHelper::OnDataAddedVerbose);
            ps_interpreter->Streams->Warning->DataAdded += gcnew EventHandler<DataAddedEventArgs ^>(h, &PSOutHelper::OnDataAddedWarning);
            ps_interpreter->Streams->Error->DataAdded += gcnew EventHandler<DataAddedEventArgs ^>(h, &PSOutHelper::OnDataAddedError);
            String ^script_code_m = gcnew String(script_code.c_str());
            ps_interpreter->AddScript(script_code_m);
            auto result = ps_interpreter->Invoke();
            for each(PSObject ^ps_object in result)
                Console::WriteLine(ps_object);
        }

    private:
        gcroot<PSOutHelper ^> h;
    };
}

int main()
{
    ps::Interpreter ps_interpreter = ps::Interpreter();
    ps_interpreter.exec_script_sync(_T(R"(Get-Command)"));
}
