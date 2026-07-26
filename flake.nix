{
  description = "NSE participant-wise OI Excel report - reproducible Python environment";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixpkgs-unstable";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, flake-utils }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = import nixpkgs { inherit system; };

        # The only external Python deps the report needs (rest is stdlib).
        pythonEnv = pkgs.python3.withPackages (ps: with ps; [
          openpyxl
          requests
        ]);
      in
      {
        # Built by run_nightly.sh via `nix build .#pythonEnv --out-link .nix-python`
        packages.pythonEnv = pythonEnv;
        packages.default = pythonEnv;

        # `nix develop` -> interactive shell with the same python
        devShells.default = pkgs.mkShell {
          packages = [ pythonEnv ];
        };

        # `nix run` -> generate the report
        apps.default = {
          type = "app";
          program = toString (pkgs.writeShellScript "participant-oi" ''
            exec ${pythonEnv}/bin/python3 ${self}/participant_oi.py "$@"
          '');
        };
      });
}
