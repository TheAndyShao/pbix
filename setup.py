import setuptools

setuptools.setup(
    # Needed to silence warnings (and to be a worthwhile package)
    name="pbix",
    author="Andrew Shao",
    author_email="andrewshao@hotmail.co.uk",
    packages=setuptools.find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=["jsonpath-ng"],
    version="1.1.2",
    license="MIT",
    description="Utilities for working with PBIX files",
)
