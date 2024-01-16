{ pkgs, ... }:

{
  # https://devenv.sh/basics/
  env.GREET = "devenv";

  # https://devenv.sh/packages/
  packages = [ pkgs.git pkgs.libyaml ];

  enterShell = ''
    echo This is the devenv shell for oddb2xml
    git --version
    ruby --version
  '';

  # https://devenv.sh/languages/
  # languages.nix.enable = true;

  languages.ruby.enable = true;
  languages.ruby.versionFile = ./.ruby-version;
  # https://devenv.sh/pre-commit-hooks/
  # pre-commit.hooks.shellcheck.enable = true;

  # https://devenv.sh/processes/
  # processes.ping.exec = "ping example.com";

  # See full reference at https://devenv.sh/reference/options/
}
