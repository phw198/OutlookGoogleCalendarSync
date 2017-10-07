# coding: utf-8

Gem::Specification.new do |spec|
  spec.name          = "jekyll-architect-theme"
  spec.version       = "0.1.0"
  spec.authors       = ["Pietro F. Menna"]
  spec.email         = ["pietromenna@yahoo.com"]

  spec.summary       = %q{Open Source version of the GitHub Pages theme, now for Jekyll}
  spec.homepage      = "https://github.com/pietromenna/jekyll-architect-theme"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").select { |f| f.match(%r{^(_layouts|_includes|_sass|LICENSE|README)/i}) }

  spec.add_development_dependency "jekyll", "~> 3.2"
  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
end
