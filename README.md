# Instant .docx Resume

```
  ╔══════════════════════════════════════════════════╗
  ║  POST /resume  →  .docx                          ║
  ║  JSON RESUME → MICROSOFT WORD                    ║
  ╚══════════════════════════════════════════════════╝
```

A minimal Node.js microservice that accepts a [JSON Resume](https://jsonresume.org/) payload and returns a formatted `.docx` file. Built with Express + `docx` npm package.

## Usage

```bash
curl -X POST https://instant-docx-resume.onrender.com/resume \
  -H "Content-Type: application/json" \
  -d @your-resume.json \
  -o Resume.docx
```

### Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `POST` | `/resume` | Accepts JSON Resume → returns `.docx` |
| `GET`  | `/health` | Health check |
| `GET`  | `/`       | Service info |

### JSON Resume Schema Adherence

```json
{
  "basics": {
    "name": "Jane Smith",
    "label": "Senior Software Engineer",
    "email": "jane@example.com",
    "phone": "555-1234",
    "url": "https://janesmith.dev",
    "location": {
      "city": "San Francisco",
      "region": "CA",
      "countryCode": "US"
    },
    "summary": "Full-stack engineer with 10+ years..."
  },
  "work": [
    {
      "company": "TechCorp",
      "position": "Senior Engineer",
      "startDate": "2020-03",
      "endDate": "2024-07",
      "summary": "Led the platform team.",
      "highlights": [
        "Designed event-driven architecture handling 50k req/s",
        "Reduced p99 latency by 40%"
      ]
    }
  ],
  "education": [
    {
      "institution": "MIT",
      "area": "Computer Science",
      "studyType": "BS",
      "startDate": "2011-09",
      "endDate": "2015-06"
    }
  ],
  "skills": [
    {
      "name": "TypeScript",
      "level": "Expert",
      "keywords": ["React", "Node.js", "Next.js"]
    }
  ],
  "projects": [
    {
      "name": "OpenAPI CLI Tool",
      "description": "Generates type-safe API clients from OpenAPI specs",
      "url": "https://github.com/jane/openapi-gen",
      "highlights": ["500+ GitHub stars"]
    }
  ],
  "certificates": [
    {
      "name": "AWS Solutions Architect",
      "issuer": "Amazon Web Services"
    }
  ],
  "awards": [
    {
      "title": "Best Engineer 2023",
      "awarder": "TechCorp"
    }
  ]
}
```
