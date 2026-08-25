# Файл: /home/detker/Документы/repository/doc-forge/shell.nix
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
