{
  description = "A very basic flake";

  outputs = { self, nixpkgs }: 
  let
    system = "x86_64-linux";
    pkgs = import nixpkgs {
      inherit system;
    };
  in
  {
    devShells.${system}.default = pkgs.mkShell {
      packages = with pkgs; [
        python311
        python311Packages.ics
        python311Packages.pytz
        python311Packages.attrs
        python311Packages.pandas
        python311Packages.openpyxl
      ];
    }; 
  };
}
