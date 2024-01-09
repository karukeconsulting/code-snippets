# Setup
The environment used when running pylint and pytest was created with miniconda, with the following command:
```batch
conda create -n code-snippets python==3.11
```

# Code description
## src/class_refactoring
A case encountered recently was the need to refactor a class method into an instance method.  
Downstream code still expected to call a class method, but the requirement was to be able to call a modified version of the method through an object instance.