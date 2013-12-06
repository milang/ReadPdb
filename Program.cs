using System;
using System.Diagnostics.SymbolStore;

using Microsoft.Samples.Debugging.CorSymbolStore;

namespace ReadPdb
{
    public static class Program
    {
        public static void Main()
        {
            var targetMethod = typeof(Program).GetMethod("Main");

            var symbolReader = SymbolAccess.GetReaderForFile(typeof (Program).Assembly.Location); // FileNotFoundException, BadImageFormatException
            var symbolMethod = symbolReader.GetMethod(new SymbolToken(targetMethod.MetadataToken));
            var docs = new ISymbolDocument[symbolMethod.SequencePointCount];
            symbolMethod.GetSequencePoints(null, docs, null, null, null, null);

            Console.WriteLine(docs[0].URL);
        }
    }
}
