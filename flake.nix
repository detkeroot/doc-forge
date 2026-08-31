# Файл: flake.nix
{
  description = "Doc-Forge: Декларативный компилятор и инспектор академических документов по ГОСТу";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, flake-utils }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = import nixpkgs { inherit system; };
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
      {
        devShells.default = pkgs.mkShell {
          name = "doc-forge-shell";
          packages = [
            pythonEnv
            pkgs.poppler-utils
            pkgs.libreoffice-qt6
          ];
          shellHook = ''
            export PYTHONPATH="$PWD:$PYTHONPATH"
            echo "🚀 Doc-Forge DevShell активен! Запуск: python3 src/cli.py --help"
          '';
        };

        packages.default = pkgs.writeShellScriptBin "doc-forge" ''
          exec ${pythonEnv}/bin/python3 ${./src/cli.py} "$@"
        '';
      }
    );
}
