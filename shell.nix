# Файл: shell.nix
{ pkgs ? import <nixpkgs> {} }:

let
  pythonEnv = pkgs.python3.withPackages (ps: with ps; [
    python-docx
    pydantic
    pyyaml
    lxml
    markdown-it-py
    jinja2
    typer
    rich
    pypdf
    pdfplumber
    pymupdf
    openpyxl
    python-pptx
    striprtf
    beautifulsoup4
    tabulate
  ]);
in
pkgs.mkShell {
  name = "doc-forge-shell";
  packages = [
    pythonEnv
    pkgs.poppler-utils
    pkgs.libreoffice-qt6
  ];
  shellHook = ''
    export PYTHONPATH="$PWD:$PYTHONPATH"
  '';
}
