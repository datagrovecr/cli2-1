using System;
using CommandLine;
using System.IO.Compression;
using Datagrove;

// folder in - defaults to current
// folder out - defaults to ../build
namespace OfficeConvert {
    
     class Options
    {
        [Option('o', "output", Required = true, HelpText = "Output directory")]
        public string? Output { get; set;}

        [Option('i', "input", Required = true, HelpText = "Input directory")]
        public string? Input { get; set;}

        [Option('r', "round", Required = false, HelpText = "Round trip directory")]
        public string? Round { get; set;}

        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; }
    }
    class Program
    {

        static async Task Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(async o =>
                {
                   await OfficeConvert.Convert(o.Input!, o.Output!, o.Round??"", false);
                });
        }


}
}


