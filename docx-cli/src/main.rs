use docx_cc;
use clap::Parser;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about=None)]
struct Args {
    #[arg(short, long)]
    template_path: String,

    #[arg(short, long, default_value_t = 1)]
    count: u8
}

fn main() {
    let args = Args::parse();

    for _ in 0..args.count {
        println!("Hello {}", args.template_path);
    }
    println!("Hello, world!");
}
