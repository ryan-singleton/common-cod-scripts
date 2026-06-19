# Common Fabrication Scripts

The Fabrication scripting syntax is rather basic and, if not handled conscientiously, can lead to some unwieldly script libraries. Keeping your scripts clean and modular is a great first step toward ensuring that future maintenance will be simple and resilient.

The contents of this script library are built from years of experience and seeing a lot of repetitive and noisy code that often looks the same. For exmaple, look at this pseudo-code.

## Shoddy example

```vb
Select item.CID 
    Case 1, 866

    If Wildcard(item.Service, "*weld*") <> 1 And 
        Wildcard(item.Filename, "square*") = 1 And 
        Wildcard(item.Material, "*galv*") = 1 Then

        If Wildcard(item.Connector[1].Value, "*TDC*") = 0 Then
            item.Connector[1].Locked = False
            item.Connector[1].Value = "TDC"
            item.Connector[1].Locked = True
            item.Update()      
        End If

        if wildcard(item.Connector[2].Value, "*TDC") = 0 True 
            item.Connector[2].Locked = flase
            item.Connector[2].Value = "TDC"
            item.Connector[1].Locked = true
            item.Update()      
        End if

        If item.CutType = "DecoledStraight" Then
            item.Dim["Lenth"].Value = "Auto"
        End If
    End If
End Select
```

There are many things sneakily wrong with this. Many points of failure.
* Using `Select` to cheat in order to do a single line for comparing the pattern number.
    * If you're using `Select` with only one `Case`, it may work, but there might be a more readable and straight forward way to do this.
* A whole lot of `Wildcard` calls with varying practices on how to check the result
    * Often the result of multiple people with different standards or habits writing chunks of the code.
* "Magic" strings. Every time we need to compare to expected text values, we explicitly write it out, with high chances of typos and failures.
* Nesting of statements with high cyclomatic complexity
* Repeated code that is should be doing the same thing, just on a different connector index.


## How these scripts help

With the scripts in this library, the functions that it provides will reduce how much code you need in the first place. On top of this, they serve as a reference for an ever-growing example script code base, so that the code you write in the future can be more stylistically and architecturally consistent.

```vb
Include MAPPATH_SCRIPTS + "Standard.cod"

dim patterns As ARRAY
patterns.Add(1, 866)

If EqualsAny(item.CID, patterns) And NotContains(item.Service, "weld") And 
    StartsWith(item.Filename, "square") And Contains(item.Material, "galv") Then

    Dim i
    For i=1 to item.Connectors
        If NotContains(item.Connector[i].Value, "tdc") Then
            ChangeItemConnector(item, i, TDC())
        End If
    Next

    If item.CutType = DecoiledStraight() Then
        item.Dim[Length()].Value = Auto()
    End If
End If
```

In the above example, there's one function that you will need to define, which is the connector function `TDC()`, which should just return the text value that your TDC connector might actually be named. Like "TDC". But outside of this, the only explicit strings written out are strings that unlikely to occur again and would hardly be considered static. They are used for ad hoc comparisons and filtering.

We use new functions that read like plain English and read legibly as to what they are actually doing. We are also employing concepts used throught the function library that aren't always prevalent in mechanical shops. For example, we're using `ARRAY` and functions that make them easier to work with.  We're using a `For` loop, and we're using consistent code styles using a VS Code extension that I discuss in the [code style](docs/CODE_STYLE.md) doc.

# Using the Scripts

I recommend pulling this repository using [git clone](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository). But you can also just copy paste the contents of the files as you like. Dealer's choice, really.

Next, I would locate the "Scripts" location for your Fabrication database. You can find this using the MAP.ini file, although it is usually somewhere like "PM Shared/Scripts".

Place the contents of "src/Scripts" from this repository into your Scripts folder.

Now, in any of your scripts, you can use these functions by adding this to your script.

`Include MAPPATH_SCRIPTS + "Standard.cod"`

From this point onward, you can use the tools found in our library. This includes within other scripts that you then call using `Include` later.

## Updating

In a command line, just go to the location where you downloaded the repository, like `cd YourDocuments/Repos/common-cod-scripts/`and run `git pull`

Now you just need to copy the contents of "src/Scripts" to your database "Scripts folder" again. If you've customized `Standard.cod`, you should not need to overwrite it.

## Customization

I suggest you do *not* edit the scripts from this repository unless you intend to contribute to this repository. The reasoning is that this library of scripts will grow over time as I and others contribute to it. If you want the updates, you will find yourself trying to stitch your changes into the updates when you pull them down. That's gonna suck.

Instead, I suggest making a new folder beside "Common" called "Custom". Inside of "Custom", you could make files like `CommonConnectors.cod` for example, that will look similar to the other `Common*.cod` files we have. But they will match your database standards for connectors.

``vb
Function TDC()
    Return "TDC"
End Function

Function Ductmate()
    Return "Ductmate 35"
End Function
``

As an example.

Now, you only need to edit `Scripts/Standard.cod`.

```vb
REM FabFunctions includes CoreFunctions, which includes the common strings.
REM Then it includes all of the remaining common functions.
Include MAPPATH_SCRIPTS "Common\FabFunctions.cod"

Include MAPPATH_SCRIPTS "Custom\CommonConnectors.cod"
Include MAPPATH_SCRIPTS "Custom\CommonSeams.cod"
```

# Language Quirks

## Functions vs Variables

The COD scripting language is a shallow derivation from VB SCript, but it lacks much of what is available there, and it carries some really bizarre rules. For example:

```vb
Function Bar()
    Return "Bar"
End Function

Dim foo = "Foo"

Function LowerCaseFooPlus(appendedValue)
    Return Lower(foo) + " " + Lower(appendedValue)
End Function
```

In object oriented languages, you can often swing code like this. But this is more of a functional programming language with some static definitions like `Job` and `item` built in. Because of this, the body of the function is only aware of two things as far as references go. The arguments passed in, such as `appendedValue` here, and other functions, such as `Bar()` in this case. The body of the function has no idea what `foo` is, because it is defined outside of the function. If you declared that variable inside of it, sure, that would work. But that won't help us consolidate. At all.

Functions are aware of other functions, however.

```vb
Function Bar()
    Return "Bar"
End Function

Function Foo()
    Return "Foo"
End Function

Function LowerCaseFooPlus(appendedValue)
    Return Lower(Foo()) + " " + Lower(appendedValue)
End Function
```

This will work. However, also note that the `Foo()` function came before `LowerCaseFooPlus()`. A function is only aware of functions defined before themselves, but that's better than not knowing of the others at all. It's weird, but that's COD scripting. Everything resembles something logical without actually being logical. I just work here.

Knowing all of this, it should be more clear what the purposes of all the `Common*.cod` files are for. This helps us to say "This text value that we rely upon throughout the code base, like `'Add Allowance To Body'`, is defined one time in one function. We use that to reference this thing, that way it will always be consistent and I won't spend hours wondering why a script didn't work because somewhere I typed `'Add AllowanceTo Body'`". We just call `AddAllowanceToBody()` always.

## Equality Testing

When testing string value equality in this language, case sensitivity is not relevant. Most people learn quickly that the code syntax itself is not case sensitive.

```vb

if true then
end if

REM works just the same as

If True Then
End IF
```

But the actual values of strings ignore case as well.

```vb
REM this is true
"TEST" = "test"

REM also true
Wildcard("TEST VALUE", "test*")
```

I discovered this long after building a bunch of utilities that ensure case could be ignored more easily. It was only when I started testing if the case-sensitive path was working that I discovered you simply cannot do it in this scripting language. Long and short of it? Your string values don't care, so don't worry about those. As for code style, see [Code Style](docs/CODE_STYLE.md).


## Contexts

The context object is something that is intended to be created once, at the start of a process, and it is used to hold state. It is optional. It has information about the user, when the process started, and can be made to hold things like logs and validation information.

After you have included the standard script, you will be able to make this call at the start of your scripted routines.

`Dim context = CreateContext("MyProcessName")`

Do this only once at the start of a routine. Don't call it in reused scripts. You want it to happen exactly once per routine. Use a pascal case string for the process name, preferably.

When you do this, the result is an `ARRAY` object in the `context`. The first member is a metadata `ARRAY` itself. Within this object, the first value will be the context you passed in ("MyProcessName"). The second will be the subcontext, which you have the option to set later. The third is the user name on the host windows machine. The fourth is a timestamp, and the fifth is a filename-compatible timestamp.

In order to make it possible for the context to know things like the username, for example, a dynamic script was generated in the windows temp folder. So the context object has this information, but a few new functions exist as well that you can use.

```vb
UserName()
or 
GetUserName(context)

Timestamp()
or
GetTimestamp(context)

FileTimestamp()
or
GetFileTimestamp(context)
```

I recommend getting it from the context if you can help it, because that is set at the beginning of the process and is now essentially a cache. But it doesn't matter much.

Keep in mind, however, that you can now add whatever you want to the context object so that you can ship around more state. For example, logs.

```vb
Dim context = CreateContext("MyProcessName")
Dim logs as ARRAY

context.Add(logs)

Function GetLogs(context)
    Return context[2]
End Functions

Function LogInfo(severity, message, context)
    Dim parts As ARRAY
    parts.Add(InBrackets(GetTimestamp(context)), InBrackets(GetUserName(context)), InBrackets(severity), message)
    Dim log = StringJoin(parts, " ")

    Dim logs = GetLogs(context)
    logs.Add(log)
End Function

Function WriteLogs(context, openAfter)
    Dim logs = GetLogs(context)

    REM you must define this in your custom stuff
    Dim logFile = GetLogFile(context)

    If logs.Count > 0 Then
        AppendArrayToFile(logFile, logs)

        If openAfter Then
            RunPath(logFile)
        End If
    End If
End Function
```