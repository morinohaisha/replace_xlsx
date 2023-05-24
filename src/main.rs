use clap::Parser;

#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
struct Args {
    #[arg(short, long)]
    template: String,

    #[arg(short, long, default_value = "out.xlsx")]
    out: String,
}

fn main() {
    let args: Args = Args::parse();
    let template: &str = args.template.as_str();
    let output: &str = args.out.as_str();
    let inputs = replace_xlsx::get_inputs().expect("Faild to input json error");

    let _ = replace_xlsx::exec_replace(template, inputs, output);
}
