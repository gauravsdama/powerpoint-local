# Third-party references

PowerPoint Local contains first-party Python source and has no third-party runtime package dependencies.

Two repositories were reviewed locally during design work but are not copied into this repository and are not runtime dependencies:

- [jongalloway/pptx-tools](https://github.com/jongalloway/pptx-tools), commit `1ff4fd10e41a9bb33853eb7b8ca5a61675571673`, MIT License, copyright Jon Galloway.
- [GongRzhe/Office-PowerPoint-MCP-Server](https://github.com/GongRzhe/Office-PowerPoint-MCP-Server), commit `3631ba2ec0c24504476f78bf74d329c9be11caaa`, MIT License, copyright GongRzhe.

The ignored `upstream/` directory is a local study checkout. It must not be included in a PowerPoint Local source or binary release.

Microsoft PowerPoint is separate proprietary software supplied by Microsoft. Mermaid Studio is a separate local application invoked through its MCP interface; it is not bundled here.
