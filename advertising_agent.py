"""
Advertising Agent for Court Reporting Businesses
Powered by Claude Opus 4.7 with adaptive thinking and agentic tool use.

This module provides an AI-powered advertising agent that helps court reporting
businesses create professional marketing content, campaign strategies, and
advertising materials targeted at law firms and legal professionals.
"""

import anthropic
import json
from typing import Generator, Optional, Dict, Any

# ---------------------------------------------------------------------------
# Model configuration
# ---------------------------------------------------------------------------
ADVERTISING_MODEL = "claude-opus-4-7"

# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------
ADVERTISING_TOOLS = [
    {
        "name": "analyze_audience",
        "description": (
            "Analyze and define the target audience for court reporting services. "
            "Returns detailed audience personas, pain points, recommended messaging, "
            "and the best marketing channels to reach each segment."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "service_type": {
                    "type": "string",
                    "description": (
                        "The specific court reporting service to analyze "
                        "(e.g., 'deposition services', 'real-time reporting', "
                        "'legal transcription', 'remote depositions', 'legal videography')"
                    ),
                },
                "geographic_focus": {
                    "type": "string",
                    "description": (
                        "Geographic area or market (e.g., 'San Pedro, Belize', "
                        "'local market', 'regional', 'national')"
                    ),
                },
                "target_firm_size": {
                    "type": "string",
                    "enum": ["solo", "small", "medium", "large", "corporate", "all"],
                    "description": "Size of law firms to target",
                },
            },
            "required": ["service_type"],
        },
    },
    {
        "name": "create_campaign_strategy",
        "description": (
            "Create a comprehensive marketing campaign strategy with phased timeline, "
            "content calendar, channel recommendations, budget allocation guidance, "
            "quick-win actions, and measurable KPIs."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "goal": {
                    "type": "string",
                    "description": (
                        "Primary campaign goal (e.g., 'attract new law firm clients', "
                        "'promote real-time reporting service', 'increase brand awareness', "
                        "'generate referrals from existing clients')"
                    ),
                },
                "duration_weeks": {
                    "type": "integer",
                    "description": "Campaign duration in weeks (1–52)",
                    "minimum": 1,
                    "maximum": 52,
                },
                "budget": {
                    "type": "string",
                    "description": (
                        "Available budget (e.g., 'minimal / under $500', "
                        "'$500–$2 000/month', '$2 000–$5 000/month', 'flexible')"
                    ),
                },
                "channels": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Preferred marketing channels to include in the plan",
                },
            },
            "required": ["goal", "duration_weeks"],
        },
    },
    {
        "name": "generate_content_brief",
        "description": (
            "Generate a structured content brief for specific advertising materials. "
            "The brief includes platform specifications, required elements, content angles, "
            "suggested keywords, and what to avoid — so the final copy is polished and effective."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content_type": {
                    "type": "string",
                    "enum": [
                        "social_media_post",
                        "email_campaign",
                        "blog_article",
                        "google_ad_copy",
                        "press_release",
                        "service_brochure",
                        "website_copy",
                        "video_script",
                        "case_study",
                        "testimonial_request",
                    ],
                    "description": "Type of advertising content to brief",
                },
                "platform": {
                    "type": "string",
                    "description": (
                        "Target platform (e.g., 'LinkedIn', 'Facebook', 'Twitter/X', "
                        "'Email Newsletter', 'Website', 'Google Ads')"
                    ),
                },
                "key_message": {
                    "type": "string",
                    "description": "Core message or value proposition to communicate",
                },
                "target_audience": {
                    "type": "string",
                    "description": "Specific audience segment to target",
                },
                "tone": {
                    "type": "string",
                    "enum": ["professional", "authoritative", "friendly", "urgent", "educational"],
                    "description": "Desired tone of the content",
                },
            },
            "required": ["content_type", "platform", "key_message"],
        },
    },
    {
        "name": "generate_ab_test_plan",
        "description": (
            "Create an A/B testing plan for advertising content including testing "
            "methodology, metrics to track, and guidance on which element to vary "
            "across multiple content variations."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content_to_test": {
                    "type": "string",
                    "description": "The original content or concept to create variations of",
                },
                "test_element": {
                    "type": "string",
                    "enum": [
                        "headline",
                        "call_to_action",
                        "value_proposition",
                        "tone",
                        "format",
                    ],
                    "description": "Which element to vary across the versions",
                },
                "num_variations": {
                    "type": "integer",
                    "description": "Number of variations to create (2–4)",
                    "minimum": 2,
                    "maximum": 4,
                },
            },
            "required": ["content_to_test", "test_element", "num_variations"],
        },
    },
]

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """You are an expert marketing and advertising agent specializing in \
court reporting and legal transcription businesses. Your deep expertise covers:

• B2B marketing to law firms, attorneys, and corporate legal departments
• Legal industry standards, tone, and professional communication
• Multi-channel digital marketing: LinkedIn, email, content marketing, Google Ads
• Brand positioning for professional services firms
• Ready-to-publish copy that legal professionals trust and respond to

Your job is to help court reporting businesses attract more clients, build their \
brand, and grow sustainably through effective, dignified advertising.

**Workflow rules:**
1. ALWAYS start by calling analyze_audience so you understand who you are addressing \
   before writing a single word of content.
2. For any campaign request, call create_campaign_strategy to build a structured plan first.
3. Before producing polished copy, call generate_content_brief to lock down the specs.
4. When A/B testing is needed, call generate_ab_test_plan for testing methodology.
5. After using the tools, provide COMPLETE, ready-to-use deliverables — not outlines. \
   Give the client something they can publish today.
6. Use a professional tone fitting the legal industry at all times.
7. Be specific, actionable, and practical — avoid generic marketing clichés."""


# ---------------------------------------------------------------------------
# Tool execution
# ---------------------------------------------------------------------------

def execute_tool(tool_name: str, tool_input: Dict[str, Any]) -> str:
    """Execute an advertising-agent tool and return a JSON string result."""

    # ------------------------------------------------------------------
    if tool_name == "analyze_audience":
        service_type = tool_input.get("service_type", "court reporting services")
        geographic_focus = tool_input.get("geographic_focus", "local market")
        target_firm_size = tool_input.get("target_firm_size", "all")

        result = {
            "service": service_type,
            "geographic_focus": geographic_focus,
            "target_firm_size": target_firm_size,
            "primary_personas": [
                {
                    "persona": "Litigation Attorney",
                    "role": "Primary buyer and end-user",
                    "pain_points": [
                        "Transcript accuracy — errors can cost cases",
                        "Tight transcript turnaround windows",
                        "Reliable scheduling around court dates",
                        "Real-time feed for fast-paced depositions",
                    ],
                    "motivations": ["Win cases", "Meet deadlines", "Reduce admin friction"],
                    "best_channels": [
                        "LinkedIn (professional posts & direct messages)",
                        "Bar association newsletters",
                        "Referrals from colleagues",
                        "Email outreach with case-study proof points",
                    ],
                },
                {
                    "persona": "Law Firm Office Manager",
                    "role": "Vendor selection & scheduling decision-maker",
                    "pain_points": [
                        "Managing multiple vendor relationships",
                        "Keeping costs under control",
                        "Avoiding last-minute scheduling crises",
                    ],
                    "motivations": ["Efficiency", "Cost predictability", "Smooth operations"],
                    "best_channels": [
                        "Email marketing (practical, ROI-focused)",
                        "Local business networking",
                        "LinkedIn",
                    ],
                },
                {
                    "persona": "Corporate / In-House Counsel",
                    "role": "Compliance-conscious buyer",
                    "pain_points": [
                        "Confidentiality and chain-of-custody assurance",
                        "Seamless technology integration",
                        "Multi-jurisdiction and remote capability",
                    ],
                    "motivations": ["Compliance", "Operational efficiency", "Cost control"],
                    "best_channels": [
                        "LinkedIn",
                        "Legal industry publications",
                        "Direct mail / targeted email",
                    ],
                },
            ],
            "key_value_propositions": [
                "Certified accuracy — every transcript guaranteed",
                "24–48-hour standard turnaround; rush options available",
                "Real-time reporting technology for complex depositions",
                "Fully remote and in-person capability",
                "Strict confidentiality and secure file delivery",
            ],
            "recommended_channels": {
                "primary": [
                    "LinkedIn (organic content + direct outreach)",
                    "Direct email to law firm contacts",
                    "Bar association member directories & publications",
                ],
                "secondary": [
                    "Google My Business (local search)",
                    "Google Search Ads (local keywords)",
                    "Legal industry directories",
                    "Client referral programme",
                ],
                "supplementary": [
                    "Facebook (community credibility)",
                    "Legal networking events / CLE sponsorships",
                    "Website SEO / blog content",
                ],
            },
            "messaging_themes": [
                "Precision and accuracy in every word",
                "Technology-forward, yet personal service",
                "Your trusted legal documentation partner",
                "We save attorney time and prevent costly errors",
                "Seamless integration with your legal workflow",
            ],
        }
        return json.dumps(result, indent=2)

    # ------------------------------------------------------------------
    elif tool_name == "create_campaign_strategy":
        goal = tool_input.get("goal", "attract new clients")
        duration_weeks = tool_input.get("duration_weeks", 8)
        budget = tool_input.get("budget", "flexible")
        channels = tool_input.get("channels", ["LinkedIn", "Email", "Google"])

        p1_end = max(2, duration_weeks // 4)
        p2_end = max(p1_end + 1, duration_weeks * 3 // 4)

        result = {
            "campaign_overview": {
                "goal": goal,
                "duration": f"{duration_weeks} weeks",
                "budget_guidance": budget,
                "primary_channels": channels,
            },
            "phases": [
                {
                    "name": f"Phase 1 — Foundation (Weeks 1–{p1_end})",
                    "focus": "Brand audit, asset creation, and list building",
                    "activities": [
                        "Audit and optimise Google My Business profile",
                        "Update/create professional LinkedIn company page",
                        "Develop core brand messaging and visual assets",
                        "Build a prospect email list (bar association directories, LinkedIn)",
                        "Create a content library of 10–15 reusable pieces",
                    ],
                    "budget_allocation": "20%",
                },
                {
                    "name": f"Phase 2 — Active Outreach (Weeks {p1_end + 1}–{p2_end})",
                    "focus": "Content publishing, paid ads, and direct outreach",
                    "activities": [
                        "Publish 3–4 social media posts per week",
                        "Launch 2 email newsletter campaigns",
                        "Run targeted LinkedIn / Google local ads",
                        "Conduct personalised outreach to 50+ prospects",
                        "Attend 1–2 bar association or legal networking events",
                    ],
                    "budget_allocation": "60%",
                },
                {
                    "name": f"Phase 3 — Optimise & Scale (Weeks {p2_end + 1}–{duration_weeks})",
                    "focus": "Performance analysis, testimonials, and retention",
                    "activities": [
                        "Analyse campaign metrics; double down on top performers",
                        "Collect and publish 3–5 client testimonials / case studies",
                        "Launch a referral incentive programme",
                        "Plan the next campaign cycle",
                    ],
                    "budget_allocation": "20%",
                },
            ],
            "content_calendar_cadence": {
                "LinkedIn": "3–4 posts per week (tips, case studies, behind-the-scenes)",
                "Email": "2 newsletters per month",
                "Blog / Website": "2 SEO-focused articles per month",
                "Google Ads": "Continuous, with weekly bid and copy optimisation",
            },
            "kpis": [
                "New client enquiries per month",
                "Website traffic increase (%)",
                "Email open rate (target ≥25 %)",
                "LinkedIn engagement rate (target ≥3 %)",
                "Lead-to-client conversion rate",
                "Cost per qualified lead",
                "Revenue generated from new clients",
            ],
            "quick_wins_this_week": [
                "Optimise Google My Business profile with photos and updated hours",
                "Ask 5 satisfied clients for a written testimonial",
                "Connect with 25 attorneys on LinkedIn with a personalised note",
                "Publish one highly-shareable educational post about what court reporters do",
            ],
        }
        return json.dumps(result, indent=2)

    # ------------------------------------------------------------------
    elif tool_name == "generate_content_brief":
        content_type = tool_input.get("content_type", "social_media_post")
        platform = tool_input.get("platform", "LinkedIn")
        key_message = tool_input.get("key_message", "")
        target_audience = tool_input.get("target_audience", "litigation attorneys")
        tone = tool_input.get("tone", "professional")

        platform_specs: Dict[str, Any] = {
            "LinkedIn": {
                "max_chars": 3000,
                "ideal_chars": "1 200–1 500",
                "format": "Text + optional image or document carousel",
                "hashtags": "3–5 relevant hashtags",
                "best_time": "Tue–Thu, 7–9 am or 5–6 pm",
            },
            "Twitter/X": {
                "max_chars": 280,
                "ideal_chars": "240 (leaves room for replies)",
                "format": "Concise text; or thread for longer content",
                "hashtags": "1–2 max",
            },
            "Facebook": {
                "ideal_chars": "40–80 (short posts perform better)",
                "format": "Text with image or short video",
                "hashtags": "0–2",
            },
            "Email Newsletter": {
                "subject_line_chars": "≤50 (preview up to 90)",
                "body": "Scannable: short paragraphs, bullet points, one clear CTA button",
                "ideal_length": "200–400 words",
                "send_day": "Tuesday or Thursday, 10 am local time",
            },
            "Google Ads": {
                "headlines": "Up to 15 (30 chars each; 3 shown per ad)",
                "descriptions": "Up to 4 (90 chars each; 2 shown per ad)",
                "display_url": "Two optional path fields, 15 chars each",
                "tips": "Include the keyword and a strong CTA in at least one headline",
            },
        }

        specs = platform_specs.get(platform, {"note": "Follow current platform best practices"})

        result = {
            "content_brief": {
                "type": content_type,
                "platform": platform,
                "target_audience": target_audience,
                "core_message": key_message,
                "tone": tone,
            },
            "platform_specifications": specs,
            "required_elements": [
                "Attention-grabbing hook or headline within the first line",
                "Core value proposition tied to the key message",
                "One clear call-to-action (e.g., 'Schedule a free consultation', 'Reply to this email')",
                "Credibility signal (years of experience, certification, client count)",
            ],
            "five_content_angles": [
                f"Problem/Solution — How {key_message} eliminates [attorney pain point]",
                f"Social Proof — A real client story that proves {key_message}",
                f"Educational — What {target_audience} should know about {key_message}",
                f"Behind-the-Scenes — How your team delivers {key_message}",
                f"Data-led — A surprising statistic that underscores the need for {key_message}",
            ],
            "power_keywords": [
                "certified court reporter",
                "accurate legal transcripts",
                "real-time deposition reporting",
                "fast turnaround",
                "confidential and secure",
            ],
            "avoid": [
                "Vague promises without specifics ('best in class', 'unmatched quality')",
                "Multiple competing calls-to-action in one piece",
                "Legal jargon your prospects may not recognise",
                "Overly salesy language — legal professionals value trust over hype",
            ],
        }
        return json.dumps(result, indent=2)

    # ------------------------------------------------------------------
    elif tool_name == "generate_ab_test_plan":
        content_to_test = tool_input.get("content_to_test", "")
        test_element = tool_input.get("test_element", "headline")
        num_variations = tool_input.get("num_variations", 3)

        element_guidance = {
            "headline": "Test different opening hooks and value statements",
            "call_to_action": "Test CTAs: 'Schedule a Consultation' vs 'Get a Free Quote' vs 'Learn More'",
            "value_proposition": "Test different core benefits: speed vs accuracy vs technology",
            "tone": "Test professional authority vs friendly approachability vs urgency",
            "format": "Test list format vs narrative paragraph vs stat-led",
        }

        result = {
            "test_overview": {
                "content_excerpt": (
                    content_to_test[:300] + "…" if len(content_to_test) > 300 else content_to_test
                ),
                "element_under_test": test_element,
                "num_variations": num_variations,
                "testing_guidance": element_guidance.get(test_element, "Test distinct approaches"),
            },
            "testing_methodology": {
                "min_run_duration": "2 weeks per variation",
                "min_impressions_per_variation": 500,
                "winner_criteria": "Higher engagement rate OR higher click-through/conversion rate",
                "statistical_confidence_target": "80 %+",
                "one_variable_rule": "Change only the tested element; keep everything else identical",
            },
            "metrics_to_track": [
                "Click-through rate (CTR)",
                "Engagement rate (reactions, shares, comments)",
                "Enquiry / form-submission conversion rate",
                "Cost per click (for paid placements)",
                "Bounce rate (for landing pages)",
            ],
            "implementation_checklist": [
                "Label each variation clearly (Control, Variant A, Variant B, …)",
                "Run Control vs one Variant at a time",
                "Record all results before moving to the next test",
                "Build a 'what worked' log for future campaigns",
            ],
        }
        return json.dumps(result, indent=2)

    # ------------------------------------------------------------------
    else:
        return json.dumps(
            {
                "error": f"Unknown tool: {tool_name}",
                "available_tools": [t["name"] for t in ADVERTISING_TOOLS],
            }
        )


# ---------------------------------------------------------------------------
# Agent runner
# ---------------------------------------------------------------------------

def run_advertising_agent(
    api_key: str,
    user_request: str,
    business_context: Optional[Dict[str, str]] = None,
) -> Generator[str, None, None]:
    """
    Run the advertising agent with an agentic tool-use loop.

    Yields:
        str — status messages (tool calls) and the final response text in markdown.

    Args:
        api_key: Anthropic API key.
        user_request: The user's advertising or marketing request.
        business_context: Optional dict with business profile information.
    """
    client = anthropic.Anthropic(api_key=api_key)

    # Build the initial user message, optionally prepended with business context
    user_content = user_request
    if business_context:
        context_lines = "\n".join(
            f"- **{k}**: {v}" for k, v in business_context.items() if v and v.strip()
        )
        if context_lines:
            user_content = (
                f"**My Court Reporting Business Profile:**\n{context_lines}\n\n"
                f"**My Request:**\n{user_request}"
            )

    messages = [{"role": "user", "content": user_content}]

    max_iterations = 12
    for iteration in range(max_iterations):
        response = client.messages.create(
            model=ADVERTISING_MODEL,
            max_tokens=8192,
            thinking={"type": "adaptive"},
            system=SYSTEM_PROMPT,
            tools=ADVERTISING_TOOLS,
            messages=messages,
        )

        # ── Final answer ────────────────────────────────────────────────
        if response.stop_reason == "end_turn":
            for block in response.content:
                if block.type == "text":
                    yield block.text
            return

        # ── Tool use ─────────────────────────────────────────────────────
        if response.stop_reason == "tool_use":
            tool_blocks = [b for b in response.content if b.type == "tool_use"]
            names = ", ".join(b.name for b in tool_blocks)
            yield f"\n\n---\n⚙️ *Running: {names}…*\n\n---\n\n"

            # Append assistant turn (includes tool_use blocks)
            messages.append({"role": "assistant", "content": response.content})

            # Execute every tool call and collect results
            tool_results = []
            for tb in tool_blocks:
                result_str = execute_tool(tb.name, tb.input)
                tool_results.append(
                    {
                        "type": "tool_result",
                        "tool_use_id": tb.id,
                        "content": result_str,
                    }
                )

            messages.append({"role": "user", "content": tool_results})
            continue

        # ── Unexpected stop reason ───────────────────────────────────────
        for block in response.content:
            if block.type == "text":
                yield block.text
        return

    yield "\n\n*[Agent reached the maximum number of iterations.]*"
