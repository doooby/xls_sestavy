# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xls_sestavy/version'

Gem::Specification.new do |spec|
  spec.name          = "xls_sestavy"
  spec.version       = XLSSestavy::VERSION
  spec.authors       = ["Ondřej Želazko"]
  spec.email         = ["zelazk.o@email.cz"]
  spec.description   = %q(Uses writeexcel and puts some helper methods on top for making summaries in xls format)
  spec.summary       = %q(xls simple helper library)
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
end
