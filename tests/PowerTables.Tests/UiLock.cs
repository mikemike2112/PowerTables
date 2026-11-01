namespace PowerTables.Tests;

/// <summary>Blocco condiviso fra le tre app per l'input reale (mouse/tastiera): solo chi lo ha usa SendInput. Cartella C:\dev\_office-ref\ui.lock.</summary>
public sealed class UiLock : IDisposable
{
    const string Path_ = @"C:\dev\_office-ref\ui.lock";
    bool _held;

    public static UiLock Acquire(int timeoutMinutes = 40)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        while (true)
        {
            try { Directory.CreateDirectory(@"C:\dev\_office-ref"); if (!Directory.Exists(Path_)) { Directory.CreateDirectory(Path_); return new UiLock { _held = true }; } }
            catch (IOException) { }
            // blocco abbandonato (più di 20 minuti): si riprende
            try { if (Directory.Exists(Path_) && DateTime.UtcNow - Directory.GetCreationTimeUtc(Path_) > TimeSpan.FromMinutes(20)) { Directory.Delete(Path_, true); continue; } } catch (IOException) { }
            if (sw.Elapsed.TotalMinutes > timeoutMinutes) throw new TimeoutException("blocco ui.lock non ottenuto");
            Thread.Sleep(3000);
        }
    }

    public void Dispose() { if (_held) { try { Directory.Delete(Path_, true); } catch (IOException) { } _held = false; } }
}
