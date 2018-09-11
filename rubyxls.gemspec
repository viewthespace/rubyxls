# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rubyxls/version'

Gem::Specification.new do |spec|
  spec.name          = "rubyxls"
  spec.version       = Rubyxls::VERSION
  spec.authors       = ["Alexander Frankel"]
  spec.email         = ["alexxander.frankel@gmail.com"]

  spec.summary       = %q{A simple DSL for generating XLS files in plain-old ruby.}
  spec.description   = %q{Generate XLS files using ruby. Rubyxls provides a simple DSL to
                          express anything from simple ruby strings to complex Active Record models.
                          Support for multi-sheet workbooks, chart generation, and formula based cells included.}
  spec.homepage      = "https://github.com/alexanderfrankel/rubyxls"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.required_ruby_version = '>= 2.2.4'

  spec.add_development_dependency "bundler"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "rspec", "~> 3.0"

  spec.add_runtime_dependency "axlsx", ">= 2.0"

end
