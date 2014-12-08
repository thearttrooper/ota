// SimpleTimer.cs
//
// Based on:
// http://freshclickmedia.co.uk/2013/11/net-stopwatch-meet-idisposable/
//

using System;
using System.Diagnostics;

class SimpleTimer : IDisposable
{
   private readonly Stopwatch _stopwatch;
   private readonly Action<Stopwatch> _action;

   public SimpleTimer(Action<Stopwatch> action = null)
   {
      _action = action ?? (s => Console.WriteLine(s.ElapsedMilliseconds));
      _stopwatch = new Stopwatch();
      _stopwatch.Start();
   }

   public void Dispose()
   {
      _stopwatch.Stop();
      _action(_stopwatch);
   }

   public Stopwatch Watch
   {
      get { return _stopwatch; }
   }
}
