# GitHub Copilot Instructions for Python Project

## Goal

Generate high-quality, efficient, maintainable, and robust Python code following modern best practices. Assume development is done using Visual Studio Code on Windows.

## General Guidelines

* **Language:** Python 3.10+
* **Style:** Adhere strictly to PEP 8. Use `black` for formatting and `ruff` or `flake8` for linting.
* **Clarity:** Prioritize readability and maintainability. Use clear variable/function names and add concise comments where logic isn't obvious.
* **Type Hinting:** Use type hints for all function signatures and complex variable assignments. Utilize the `typing` module.
* **Docstrings:** Generate Google-style docstrings for all modules, classes, functions, and methods. Include `Args:`, `Returns:`, and `Raises:`.

## Modern Python Features

* **Modern Features:**
  * Pattern matching (Python 3.10+)
  * Structural pattern matching (match/case)
  * The walrus operator (:=) for assignment expressions
  * Dataclasses or Pydantic models for data representation
  * F-strings with inline formatting (Python 3.8+)
  * Type unions with pipe operator (Python 3.10+)
  * Improved type annotations (TypeAlias, TypeGuard, ParamSpec)
  * Self type for recursive type hints (Python 3.11+)

## Code Generation

* **Optimization:**
    * Prefer built-in functions and standard library modules where possible.
    * Use list comprehensions, generator expressions, and dictionary comprehensions for concise and efficient iterations.
    * Be mindful of algorithmic complexity (Big O notation). Suggest efficient algorithms for the task.
    * Avoid premature optimization; focus on clarity first unless performance is critical.
* **Accuracy:**
    * Pay close attention to the surrounding code context and comments.
    * If requirements are ambiguous, generate a simple, clear implementation and add a comment suggesting potential alternatives or asking for clarification.
    * Generate code that directly addresses the prompt or comment preceding it.

## Directory Structure

Assume and maintain the following project structure:

```
project-root/
├── .github/
│   └── copilot-instructions.md
├── src/
│   └── sf132_sf133_recon/
│       ├── __init__.py
│       ├── main.py
│       ├── core/
│       ├── modules/
│       └── utils/
├── tests/
│   ├── __init__.py
│   └── test_*.py
├── docs/
│   └── ...
├── data/          # Optional: For data files
├── scripts/       # Optional: For helper scripts
├── .gitignore
├── pyproject.toml # Or requirements.txt
└── README.md
```

* Place core application logic within `src/sf132_sf133_recon/`.
* Place unit and integration tests within `tests/`. Test files should mirror the structure of the `src/` directory.
* Use relative imports within the `src/` directory (e.g., `from .core import ...`).

## Dependency Management

* **Dependency Management:**
  * Use Poetry for comprehensive dependency management
  * Separate dev dependencies from production dependencies
  * Pin dependencies with specific versions for reproducibility
  * Consider using a tool like `pip-compile` for generating deterministic requirements.txt
  * Use `pipdeptree` to visualize and manage complex dependency trees
  * Implement dependency security scanning in CI pipeline

## Virtual Environment Management

* **Environment Management:**
  * Use venv or virtualenv for isolated environments
  * Consider using pyenv for Python version management
  * Document environment setup steps in README
  * Create environment setup scripts if complex
  * Use .env files with python-dotenv for environment variables
  * Consider creating a dev container configuration for VSCode

## Error Handling

* **Error Handling Strategy:**
  * Catch specific exceptions rather than generic `Exception`.
  * Define custom exception classes for application-specific errors when appropriate.
  * Use `try...except...finally` blocks for cleanup operations (e.g., closing files or network connections).
  * Use context managers (`with` statement) for resource management (files, locks, connections).
  * Provide clear and informative error messages.
  * Design an error hierarchy specific to application domains
  * Implement error codes for systematic troubleshooting
  * Create appropriate recovery mechanisms for different error types
  * Log contextual information with exceptions
  * Consider using exception chaining with `raise ... from ...` syntax

## Debugging

* **Variable Names:** Use descriptive variable names.
* **Intermediate Variables:** Don't shy away from using intermediate variables to clarify steps in complex calculations or logic.
* **Pure Functions:** Prefer pure functions (functions whose output depends only on their input and have no side effects) where possible, as they are easier to test and debug.
* **Assertions:** Use `assert` statements for sanity checks during development (but be aware they can be disabled).
* **Debugging Tools:**
  * Utilize VSCode's built-in debugger with breakpoints
  * Consider using pdb/ipdb for command-line debugging
  * Use logging for persistent debugging information
  * Add debugging decorators for function entry/exit tracing
  * Implement custom debug views for complex data structures

## Logging

* **Standard Library:** Use the built-in `logging` module.
* **Configuration:** Configure logging early in the application's entry point (`main.py`). Consider configuration via file (`logging.conf`) or dictionary (`logging.config.dictConfig`).
* **Levels:** Use appropriate logging levels (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`).
* **Context:** Include relevant context in log messages (e.g., function names, relevant variable values).
* **Avoid `print()`:** Replace `print()` statements used for debugging or status updates with appropriate `logging` calls.
* **Advanced Logging:**
  * Implement structured logging with JSON formatter for machine parsing
  * Use log rotation to manage log file sizes
  * Consider centralized logging for distributed applications
  * Add correlation IDs for tracking requests across components
  * Use custom log handlers for specific log destinations

## Testing Strategy

* **Testing Framework:**
  * Use `pytest` as the testing framework.
  * Generate tests that aim for high code coverage (80%+).
  * Utilize `pytest` fixtures for setting up and tearing down test states.
  * Use `unittest.mock` (or `pytest-mock`) for mocking dependencies.
  * Test functions should start with `test_`.
* **Comprehensive Testing:**
  * Unit tests for individual functions
  * Integration tests for component interactions
  * End-to-end tests for critical user flows
  * Property-based testing with hypothesis for complex logic
  * Performance tests for critical paths
  * Parameterized tests for multiple inputs
* **Test Quality:**
  * Follow AAA pattern (Arrange, Act, Assert)
  * Create focused tests with clear objectives
  * Use descriptive test names that explain the test purpose
  * Implement test coverage reporting and enforcement
  * Consider mutation testing to evaluate test effectiveness

## Documentation

* **Documentation Tools:**
  * Use Sphinx or MkDocs for generating documentation
  * Include code examples in docstrings to demonstrate usage
  * Consider using autodoc extensions to generate API docs from docstrings
  * Create a detailed README.md with installation, usage, and contribution guidelines
  * Generate online documentation on each release
* **Documentation Content:**
  * Include architecture diagrams (consider using Mermaid or PlantUML)
  * Document design decisions and trade-offs
  * Create user guides with examples
  * Maintain API documentation with examples
  * Document configuration options and environment variables

## Security

* **Security Practices:**
  * Always validate and sanitize external input (user input, API responses, file contents).
  * Do not hardcode secrets (API keys, passwords). Use environment variables or a dedicated secrets management tool.
  * Suggest placeholders like `os.getenv("API_KEY")` or use python-dotenv.
  * Be mindful of third-party library security vulnerabilities.
  * Implement input validation with pydantic or marshmallow
  * Use parameterized queries for database access
  * Conduct regular dependency scanning with safety
  * Follow OWASP guidelines for web applications
  * Implement content security policies where appropriate
  * Use secure hashing and password storage (argon2, bcrypt)

## Code Quality Automation

* **Pre-commit Hooks:**
  * Set up pre-commit hooks for:
    * Code formatting (black)
    * Import sorting (isort)
    * Linting (ruff/flake8)
    * Type checking (mypy)
    * Security scanning (bandit)
  * Include .pre-commit-config.yaml in project
* **CI/CD Integration:**
  * Set up GitHub Actions workflows for automated testing
  * Configure linting and type checking in CI pipelines
  * Implement automated deployment processes
  * Set up code coverage reporting
  * Implement parallel testing for faster feedback
  * Create deployment pipelines for different environments

## Static Type Checking

* **Type Checking Tools:**
  * Use mypy for static type checking
  * Configure mypy.ini or pyproject.toml with appropriate strictness
  * Consider runtime type validation with libraries like Pydantic
  * Use Protocol classes for duck typing
  * Implement TypeGuard for runtime type narrowing
  * Use TypeVar for generic function definitions
  * Create custom types for domain-specific concepts

## Performance Optimization

* **Performance Analysis:**
  * Profile code with cProfile or py-spy
  * Use line_profiler for line-by-line profiling
  * Consider memory_profiler for memory usage analysis
  * Implement benchmarking tests for critical paths
  * Monitor performance metrics in production
* **Optimization Techniques:**
  * Use caching strategically (functools.lru_cache, redis)
  * Implement batch processing for I/O operations
  * Consider using compiled extensions for performance-critical sections
  * Parallelize CPU-bound operations with multiprocessing
  * Optimize database queries with proper indexing

## Asynchronous Programming

* **Async Development:**
  * Use asyncio for I/O-bound operations
  * Prefer async/await syntax over callbacks
  * Consider libraries like aiohttp for async HTTP
  * Use asyncio.gather for concurrent tasks
  * Implement proper error handling in async code
  * Be mindful of event loop blocking
  * Consider using async frameworks for web applications

## Containerization

* **Container Strategy:**
  * Provide Dockerfile for consistent environment
  * Use multi-stage builds for optimized containers
  * Create docker-compose.yml for local development
  * Document container usage in README
  * Optimize container image size
  * Include health checks in container configurations
  * Implement proper signal handling for graceful shutdowns

## Cross-Platform Compatibility

* **Platform Independence:**
  * Use pathlib for file path manipulation
  * Avoid platform-specific commands or libraries
  * Test on multiple platforms when possible
  * Use CI with multiple OS runners
  * Handle line ending differences properly
  * Be mindful of file permission differences between platforms
  * Use platform-independent temporary file handling

## Project-Specific Guidance

* **Data Processing:**
  * Prefer pandas for tabular data manipulation
  * Use numpy for numerical operations
  * Consider dask for larger-than-memory datasets
  * Document data schemas and validation methods
  * Implement data quality checks and validation

* **API Development:** (if applicable)
  * Use FastAPI for modern API development
  * Design RESTful endpoints following best practices
  * Implement comprehensive request validation
  * Document APIs with OpenAPI/Swagger
  * Implement proper authentication and authorization
  * Use rate limiting for public endpoints

## Version Control Best Practices

* **Git Workflow:**
  * Use descriptive branch names (feature/, bugfix/, etc.)
  * Write clear, concise commit messages
  * Keep commits focused and atomic
  * Use Pull Requests for code review
  * Implement branch protection rules
  * Consider using Conventional Commits format
  * Use git hooks for quality assurance