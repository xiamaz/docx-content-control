use std::fs;
use std::io::{BufReader, BufWriter};
use docx_cc;
use clap::{Parser, Subcommand};

#[derive(Parser, Debug)]
#[command(author, version, about, long_about=None)]
struct Args {
    #[command(subcommand)]
    command: Commands,

    #[arg(short, long)]
    template_path: String,
}

#[derive(Debug, Subcommand)]
enum Commands {
    Clear {
        #[arg(last=true)]
        output_path: String,
    }
}

fn load_path(path: &str) -> docx_cc::ZipData {
    let fname = std::path::Path::new(&path);
    let file = fs::File::open(fname).unwrap();
    let reader = BufReader::new(file);
    docx_cc::list_zip_contents(reader).unwrap()
}

fn main() {
    let args = Args::parse();

    let data = load_path(&args.template_path);

    match args.command {
        Commands::Clear { output_path } => {
            let result = docx_cc::remove_content_controls(&data);
            let output_file = fs::File::create(output_path).unwrap();
            let mut writer = BufWriter::new(output_file);
            let _ = docx_cc::zip_dir(&result, &mut writer);
        }
    }
}
