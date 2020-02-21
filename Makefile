format: format-ruby format-python

format-ruby:
	rufo *.rb

format-python:
	python3 -m black *.py

PHONY: format format-ruby format-python
