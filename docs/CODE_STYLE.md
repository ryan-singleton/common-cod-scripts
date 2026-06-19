# Code style

When writing this code, it is good to use the [agilebim.fabcod](https://github.com/AgileBIM/FabCOD) extension while writing these scripts. It will provide help while writing your code, as long as CADmep of some version is running in the background. If you want to know more about why this is, I think it's because of something called an LSP, or language server provider. It's complicated. But if you can swing it, it helps when writing code immensely.

Because of this extension, it forces us to be a lot more consistent about how our code is written. For example, `Function` and `End Function` instead of `function` and `end function`. This will feel weird at first to a lot of us, but whatever. The tool is going to insist on it and it's best not to swim uphill, or .. wait, what?

Anyway, please remain consistent in how the code is written.

- Use `Dim i` when declaring a variable that will be used in a `For` loop please. This is a pretty standard practice in all of programming. Or `index`.
    - Naming it `Conn` or something like that implies that the variable actually contains a connector. It's weird and hard to read. You're trying to create something that is extremely easy to read, not something that adds friction.
- Use `camelCase` for variables or function arguments
- Identify tasks and subtasks, place them into reusable functions, and locate them in the appropriate files.

It is tempting to use `Select` when matching against multiuple values on a single property, but this a misuse of the syntax that ccan be much less readable later on.

```vb
Select item.CID
    Case 1,2,3
    REM do stuff
End Select
```

`Select` is meant to handle multiple cases. Not a single case against multiple values.

```vb
Select item.CID
    Case 1,2,3
    REM do one thing
    Case 4,5
    REM do other thing
End Select
```

For a single case compared against multiple values, we now have a few functions that should make it simple to do with an `ARRAY`.

```vb
Dim patterns as ARRAY
patterns.Add(1,2,3)

If EqualsAny(item.CID, patterns) Then
    REM do stuff
End If
```

This reads more clearly and uses the language as it should be used. Notice that you can add multiple members to an array at once. Or you can add multiple members on multiple lines if there are a lot to add.

```vb
Function PittsPatterns()
    Dim patterns As ARRAY
    patterns.Add(1,2,3,4,6,7,9,11,12,13,14,15,17,18,19,20,37,166,222)
    patterns.Add(328,329,330,354,382,399,522,525,761,861,866,932)
    Return patterns
End Function
```