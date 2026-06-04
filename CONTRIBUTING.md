Thanks for contributing! Please follow these guidelines when submitting changes.

1. Fork the repository and create a topic branch for your change:

   git checkout -b my-feature

2. Run the project locally and validate changes:

   source venv/bin/activate
   python -m pip install -r requirements.txt
   python main.py --dry-run

3. Commit changes with a clear message and open a Pull Request against `main`.

4. Keep PRs small and focused. Provide screenshots or logs if your change affects runtime behavior.

5. For security-sensitive issues (credential leaks, PII exposure), do NOT open a public issue — contact the repository owner directly.

Maintainers will review and merge PRs; please be responsive to review feedback.
