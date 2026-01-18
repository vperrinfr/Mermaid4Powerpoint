# Contributing to Mermaid4Powerpoint

Thank you for your interest in contributing to Mermaid4Powerpoint! This document provides guidelines for contributing to the project.

## How to Contribute

### Reporting Bugs

If you find a bug, please open an issue on GitHub with:
- A clear, descriptive title
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Your environment (OS, PowerPoint version, Node.js version)
- Screenshots if applicable

### Suggesting Enhancements

We welcome feature requests! Please open an issue with:
- A clear description of the feature
- Why this feature would be useful
- Examples of how it would work
- Any relevant mockups or diagrams

### Pull Requests

1. **Fork the repository** and create your branch from `main`
2. **Make your changes** following our coding standards
3. **Test your changes** thoroughly
4. **Update documentation** if needed
5. **Submit a pull request** with a clear description

## Development Setup

1. Clone your fork:
```bash
git clone https://github.com/YOUR_USERNAME/Mermaid4Powerpoint.git
cd Mermaid4Powerpoint
```

2. Install dependencies:
```bash
npm install
```

3. Generate SSL certificates:
```bash
npx office-addin-dev-certs install
```

4. Start development server:
```bash
npm start
```

## Coding Standards

### TypeScript
- Use TypeScript for all new code
- Enable strict mode
- Add type annotations where helpful
- Avoid `any` types when possible

### React
- Use functional components with hooks
- Keep components small and focused
- Use meaningful component and variable names
- Add comments for complex logic

### Code Style
- Use 2 spaces for indentation
- Use semicolons
- Use single quotes for strings
- Follow existing code patterns

### Commits
- Write clear, descriptive commit messages
- Use present tense ("Add feature" not "Added feature")
- Reference issues in commits when applicable

## Testing

Before submitting a PR:
- Test in PowerPoint (Windows, Mac, or Online)
- Test all diagram types (flowchart, sequence, class)
- Test error handling with invalid syntax
- Check browser console for errors
- Verify no TypeScript compilation errors

## Documentation

- Update README.md for new features
- Add JSDoc comments for functions
- Update SETUP_GUIDE.md if setup changes
- Include examples for new diagram types

## Project Structure

```
Mermaid4Powerpoint/
├── src/
│   ├── taskpane/        # Main UI components
│   └── commands/        # Office command handlers
├── assets/              # Icons and images
├── picture/             # Screenshots
├── manifest.xml         # Office Add-in manifest
└── webpack.config.js    # Build configuration
```

## Adding New Diagram Types

To add a new Mermaid diagram type:

1. Add example to `defaultDiagrams` in `src/taskpane/App.tsx`
2. Update the dropdown options
3. Test rendering and insertion
4. Update documentation with examples
5. Add to README.md features list

## Code Review Process

1. All PRs require review before merging
2. Address review comments promptly
3. Keep PRs focused on a single feature/fix
4. Ensure CI checks pass

## Community Guidelines

- Be respectful and inclusive
- Help others learn and grow
- Give constructive feedback
- Follow the [Code of Conduct](CODE_OF_CONDUCT.md)

## Questions?

- Open a GitHub issue for questions
- Check existing issues and documentation first
- Be patient - maintainers are volunteers

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

Thank you for contributing to Mermaid4Powerpoint! 🎉