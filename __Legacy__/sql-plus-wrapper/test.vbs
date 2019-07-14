x = CreateObject("WScript.Shell").Run("sqlplus DIHASSANEIN@EWH9 < tmp.txt", 0, True)

Call WScript.StdOut.WriteLine(x)
