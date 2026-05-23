"""
Advertising Agent — General-Purpose Business Marketing Assistant
Powered by Claude Opus 4.7 with adaptive thinking and agentic tool use.

Works for any business type: retail, services, tech, hospitality, healthcare, etc.
"""

import anthropic
import json
from typing import Generator, Optional, Dict, Any

ADVERTISING_MODEL = "claude-opus-4-7"

# ---------------------------------------------------------------------------
# Tool definitions  (fully generic — no industry assumptions)
# ---------------------------------------------------------------------------
ADVERTISING_TOOLS = [
    {
        "name": "analyze_audience",
        "description": (
            "Analyze and define the ideal target audience for a business. "
            "Returns detailed buyer personas, pain points, motivations, "
            "and the best channels to reach each segment."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "business_type": {
                    "type": "string",
                    "description": "What the business does (e.g., 'online clothing store', 'IT consulting firm', 'restaurant', 'SaaS startup')",
                },
                "product_or_service": {
                    "type": "string",
                    "description": "The specific product or service being advertised",
                },
                "current_audience": {
                    "type": "string",
                    "description": "Who is currently buying (if known), or leave blank to discover from scratch",
                },
                "goal": {
                    "type": "string",
                    "description": "What the business wants from this audience (e.g., 'first purchase', 'repeat orders', 'enterprise contracts')",
                },
            },
            "required": ["business_type", "product_or_service"],
        },
    },
    {
        "name": "build_campaign_strategy",
        "description": (
            "Build a full go-to-market advertising campaign strategy — "
            "with phases, channel mix, content cadence, budget guidance, "
            "quick wins, and measurable KPIs."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "business_type": {
                    "type": "string",
                    "description": "What the business does",
                },
                "campaign_objective": {
                    "type": "string",
                    "description": (
                        "Primary goal (e.g., 'launch new product', 'grow social following', "
                        "'drive website sales', 'generate B2B leads', 'increase foot traffic')"
                    ),
                },
                "duration_weeks": {
                    "type": "integer",
                    "description": "How many weeks the campaign should run (1–52)",
                    "minimum": 1,
                    "maximum": 52,
                },
                "monthly_budget": {
                    "type": "string",
                    "description": "Approximate monthly budget (e.g., 'under $500', '$1 000–$3 000', '$5 000+', 'unknown')",
                },
                "channels": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Channels to include (e.g., ['Instagram', 'Email', 'Google Ads', 'TikTok'])",
                },
            },
            "required": ["business_type", "campaign_objective", "duration_weeks"],
        },
    },
    {
        "name": "create_content_brief",
        "description": (
            "Create a structured content brief for any advertising material. "
            "Defines platform specs, required elements, messaging angles, "
            "tone guidance, and what to avoid — so the final copy is on-brand and effective."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content_format": {
                    "type": "string",
                    "enum": [
                        "social_media_post",
                        "paid_ad_copy",
                        "email_campaign",
                        "blog_article",
                        "video_script",
                        "website_copy",
                        "product_description",
                        "press_release",
                        "sms_campaign",
                        "podcast_ad_script",
                    ],
                    "description": "Type of advertising content",
                },
                "platform": {
                    "type": "string",
                    "description": "Target platform (e.g., 'Instagram', 'LinkedIn', 'Google Ads', 'Email', 'TikTok', 'Facebook', 'Twitter/X', 'YouTube')",
                },
                "core_message": {
                    "type": "string",
                    "description": "The single most important thing this content must communicate",
                },
                "target_persona": {
                    "type": "string",
                    "description": "Who this specific piece is aimed at",
                },
                "desired_action": {
                    "type": "string",
                    "description": "What you want the audience to do after seeing this (e.g., 'click Buy Now', 'sign up for a free trial', 'book a call')",
                },
                "tone": {
                    "type": "string",
                    "enum": ["professional", "casual", "playful", "urgent", "inspirational", "educational", "luxury", "bold"],
                    "description": "Brand tone for this piece",
                },
            },
            "required": ["content_format", "platform", "core_message", "desired_action"],
        },
    },
    {
        "name": "plan_ab_test",
        "description": (
            "Design an A/B testing plan for advertising content — "
            "specifying what to vary, how long to run each version, "
            "which metrics to track, and how to declare a winner."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "content_to_test": {
                    "type": "string",
                    "description": "The original content, concept, or ad to test",
                },
                "variable_to_test": {
                    "type": "string",
                    "enum": ["headline", "call_to_action", "image_or_creative", "offer", "tone", "audience_segment", "send_time"],
                    "description": "The single variable to change across versions",
                },
                "num_versions": {
                    "type": "integer",
                    "description": "Total number of versions including the control (2–5)",
                    "minimum": 2,
                    "maximum": 5,
                },
                "success_metric": {
                    "type": "string",
                    "description": "How to measure the winner (e.g., 'click-through rate', 'conversion rate', 'open rate', 'cost per lead')",
                },
            },
            "required": ["content_to_test", "variable_to_test", "num_versions"],
        },
    },
]

# ---------------------------------------------------------------------------
# System prompt — generic, works for any business
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """You are an expert advertising and marketing strategist who works with \
businesses of all types — from solo entrepreneurs to established companies, across every \
industry.

Your capabilities:
• Audience research and buyer persona development
• Multi-channel campaign strategy (paid, organic, email, content, social)
• Copywriting: ads, emails, social posts, landing pages, video scripts, blog articles
• Brand voice development and messaging frameworks
• A/B testing design and performance analysis
• Budget-conscious strategies that maximise ROI

How you work:
1. ALWAYS call analyze_audience first so you understand exactly who you are writing for.
2. For campaign requests, call build_campaign_strategy before producing any content.
3. Before writing polished copy, call create_content_brief to lock in the specs.
4. When testing is needed, use plan_ab_test to design a rigorous testing methodology.
5. After using the tools, produce COMPLETE, ready-to-use deliverables — not outlines or \
   bullet-point skeletons. Give the client something they can act on immediately.
6. Match your tone to the business and their audience — never default to generic corporate speak.
7. Be specific, creative, and practical. Avoid marketing clichés."""


# ---------------------------------------------------------------------------
# Tool execution
# ---------------------------------------------------------------------------

def execute_tool(tool_name: str, tool_input: Dict[str, Any]) -> str:
    """Execute a tool call and return a JSON string result."""

    # ── analyze_audience ─────────────────────────────────────────────────────
    if tool_name == "analyze_audience":
        business_type = tool_input.get("business_type", "business")
        product = tool_input.get("product_or_service", "product/service")
        current_audience = tool_input.get("current_audience", "")
        goal = tool_input.get("goal", "attract new customers")

        result = {
            "business": business_type,
            "product_or_service": product,
            "stated_goal": goal,
            "existing_audience_note": current_audience or "Not provided — personas built from scratch",
            "primary_personas": [
                {
                    "name": "Core Buyer",
                    "description": f"The most likely paying customer for {product}",
                    "typical_triggers": [
                        "Actively searching for a solution to a specific problem",
                        "Saw a recommendation from someone they trust",
                        "Encountered the brand multiple times and became curious",
                    ],
                    "key_objections": [
                        "Is this worth the price?",
                        "Can I trust this brand?",
                        "Does it actually work for someone like me?",
                    ],
                    "best_channels": ["Search (Google/Bing)", "Word of mouth / referrals", "Social proof & reviews"],
                },
                {
                    "name": "Aspirational Buyer",
                    "description": "Aware of the problem but not yet actively searching",
                    "typical_triggers": [
                        "Inspired by content that shows the transformation",
                        "Influenced by peers or influencers",
                        "Responds to a compelling offer or limited-time deal",
                    ],
                    "key_objections": [
                        "I'm not sure I really need this right now",
                        "I haven't heard of this brand before",
                    ],
                    "best_channels": ["Instagram / TikTok", "YouTube", "Influencer partnerships", "Retargeting ads"],
                },
                {
                    "name": "Repeat / Loyalty Buyer",
                    "description": "Past customer with potential for repeat purchase or upsell",
                    "typical_triggers": [
                        "Positive first experience",
                        "Personalised follow-up offer",
                        "Loyalty reward or exclusive access",
                    ],
                    "key_objections": [
                        "I'm already happy with what I have",
                        "Why should I buy again / upgrade?",
                    ],
                    "best_channels": ["Email marketing", "SMS", "Loyalty programme", "Exclusive offers"],
                },
            ],
            "universal_messaging_pillars": [
                "Lead with the transformation — what life looks like AFTER buying",
                "Use social proof heavily (reviews, numbers, testimonials)",
                "Reduce perceived risk (guarantee, free trial, easy returns)",
                "Create urgency without being manipulative",
                "Speak to one person, not a crowd",
            ],
            "channel_priority_framework": {
                "highest_intent": ["Google Search Ads", "SEO", "Reviews / Directories"],
                "community_building": ["Instagram", "Facebook", "LinkedIn", "TikTok"],
                "owned_channels": ["Email list", "SMS", "Blog / Content hub"],
                "amplification": ["Influencers", "Paid social", "Podcast sponsorships"],
            },
        }
        return json.dumps(result, indent=2)

    # ── build_campaign_strategy ──────────────────────────────────────────────
    elif tool_name == "build_campaign_strategy":
        business_type = tool_input.get("business_type", "business")
        objective = tool_input.get("campaign_objective", "grow the business")
        weeks = tool_input.get("duration_weeks", 8)
        budget = tool_input.get("monthly_budget", "flexible")
        channels = tool_input.get("channels", ["Social Media", "Email", "Paid Ads"])

        p1 = max(2, weeks // 4)
        p2 = max(p1 + 2, (weeks * 3) // 4)

        result = {
            "campaign_overview": {
                "business": business_type,
                "objective": objective,
                "duration": f"{weeks} weeks",
                "monthly_budget": budget,
                "channels": channels,
            },
            "phases": [
                {
                    "phase": f"Phase 1 — Build (Weeks 1–{p1})",
                    "focus": "Set the foundation before spending money on reach",
                    "activities": [
                        "Audit and optimise all existing channels and profiles",
                        "Define brand voice, key messages, and visual style",
                        "Create a core content library (10+ reusable assets)",
                        "Set up tracking: pixels, UTMs, conversion goals",
                        "Identify and brief 2–3 micro-influencers or partners if relevant",
                    ],
                    "budget_split": "15%",
                },
                {
                    "phase": f"Phase 2 — Launch & Reach (Weeks {p1 + 1}–{p2})",
                    "focus": "Push content out, run ads, and collect real data",
                    "activities": [
                        "Publish content consistently across chosen channels",
                        "Launch paid campaigns with small test budgets first",
                        "Run email welcome sequence to new subscribers",
                        "Engage actively with comments and DMs",
                        "Capture leads and nurture them toward purchase",
                    ],
                    "budget_split": "65%",
                },
                {
                    "phase": f"Phase 3 — Optimise & Scale (Weeks {p2 + 1}–{weeks})",
                    "focus": "Cut what doesn't work, scale what does",
                    "activities": [
                        "Pause underperforming ads; increase budget on winners",
                        "Collect customer testimonials and social proof",
                        "Launch retargeting campaigns to warm audiences",
                        "Plan the next campaign cycle using data from this one",
                    ],
                    "budget_split": "20%",
                },
            ],
            "content_cadence": {
                "Social Media (organic)": "4–5 posts per week",
                "Email": "1–2 campaigns per week",
                "Blog / SEO content": "1–2 articles per month",
                "Paid ads": "Always-on, optimised weekly",
                "Video (if applicable)": "1–2 pieces per month",
            },
            "kpis": [
                "Cost per lead (CPL)",
                "Cost per acquisition (CPA)",
                "Return on ad spend (ROAS)",
                "Email open rate (target: 25%+) and click rate (target: 3%+)",
                "Social engagement rate (target: 3–5%)",
                "Website conversion rate",
                "Revenue directly attributed to campaign",
            ],
            "quick_wins_start_today": [
                "Post one piece of customer-proof content (review, result, before/after)",
                "Set up or optimise your Google Business Profile",
                "Email your existing contact list with a genuine update and offer",
                "Identify your single best-performing past post and boost it with $20",
            ],
        }
        return json.dumps(result, indent=2)

    # ── create_content_brief ─────────────────────────────────────────────────
    elif tool_name == "create_content_brief":
        fmt = tool_input.get("content_format", "social_media_post")
        platform = tool_input.get("platform", "Instagram")
        core_message = tool_input.get("core_message", "")
        persona = tool_input.get("target_persona", "ideal customer")
        action = tool_input.get("desired_action", "visit our website")
        tone = tool_input.get("tone", "casual")

        specs: Dict[str, Any] = {
            "Instagram": {"caption_chars": "≤2 200 (first 125 chars are critical)", "hashtags": "5–10", "visual": "High-quality image or Reel", "best_time": "Tue–Fri, 11am–1pm or 7–9pm"},
            "TikTok": {"video_length": "15–60 sec sweet spot", "hook": "First 2 seconds must stop the scroll", "caption": "≤150 chars", "hashtags": "3–5 trending + niche mix"},
            "Facebook": {"post_chars": "40–80 chars outperform longer posts", "image": "1 200×630 px", "hashtags": "1–2 max", "best_time": "Wed, 11am–1pm"},
            "LinkedIn": {"post_chars": "1 300 chars ideal", "hook": "First line must earn the 'see more' click", "hashtags": "3–5 professional", "best_time": "Tue–Thu, 8–10am"},
            "Twitter/X": {"char_limit": "280 (aim for ≤240)", "threads": "Use for long-form; hook tweet is critical", "hashtags": "1–2 max"},
            "Google Ads": {"headlines": "Up to 15 @ 30 chars each (3 shown)", "descriptions": "Up to 4 @ 90 chars each (2 shown)", "tip": "Put keyword + CTA in at least one headline"},
            "Email": {"subject": "≤50 chars; avoid spam trigger words", "preview_text": "90 chars", "body": "One idea, one CTA; bullet points for scannability", "best_day": "Tue or Thu, 10am"},
            "YouTube": {"title": "≤60 chars, keyword-first", "description": "First 2–3 lines show without expanding", "hook": "First 30 sec must justify watching"},
        }

        platform_spec = specs.get(platform, {"note": "Follow current platform best practices for format and character limits"})

        result = {
            "brief": {
                "format": fmt,
                "platform": platform,
                "target_persona": persona,
                "core_message": core_message,
                "desired_action": action,
                "tone": tone,
            },
            "platform_specs": platform_spec,
            "must_have_elements": [
                "Pattern-interrupt hook in the very first line / frame",
                f"Core message: '{core_message}'",
                f"One clear CTA: '{action}'",
                "Social proof element (number, review snippet, client name)",
                "Brand voice consistency with the '{tone}' tone",
            ],
            "proven_content_angles": [
                f"Before/After — show the transformation your product/service creates",
                f"Problem/Agitate/Solve — name the pain, make it vivid, offer the fix",
                f"Social proof story — a real customer result tied to '{core_message}'",
                f"Curiosity gap — tease insight that {persona} doesn't want to miss",
                f"Contrast — what life looks like without vs. with your solution",
            ],
            "tone_guidance": {
                "professional": "Authoritative sentences, data-backed claims, no slang",
                "casual": "Conversational, first-person, contractions OK, feel like a friend",
                "playful": "Puns/wordplay welcome, emoji OK, light and energetic",
                "urgent": "Time-bound language, scarcity signals, action verbs",
                "inspirational": "Aspirational vision, 'you can do this' energy, emotional",
                "educational": "Teach something valuable, step-by-step, 'did you know' format",
                "luxury": "Understated confidence, sensory language, exclusivity cues",
                "bold": "Short punchy sentences. Strong claims. No hedging.",
            }.get(tone, "Match the brand's natural voice"),
            "avoid": [
                "Multiple calls-to-action competing with each other",
                "Opening with 'We' — start with 'You' or a hook",
                "Vague claims without specifics ('amazing quality', 'best service')",
                "Jargon your audience wouldn't use themselves",
            ],
        }
        return json.dumps(result, indent=2)

    # ── plan_ab_test ─────────────────────────────────────────────────────────
    elif tool_name == "plan_ab_test":
        content = tool_input.get("content_to_test", "")
        variable = tool_input.get("variable_to_test", "headline")
        versions = tool_input.get("num_versions", 2)
        metric = tool_input.get("success_metric", "click-through rate")

        guidance = {
            "headline": "Write versions with different emotional angles (curiosity vs. urgency vs. benefit-led)",
            "call_to_action": "Test button/link text: action verb + outcome (e.g., 'Get My Free Quote' vs. 'Start Saving Today' vs. 'Book a Call')",
            "image_or_creative": "Test lifestyle image vs. product-only vs. person-facing-camera vs. text-overlay graphic",
            "offer": "Test price framing (monthly vs. annual), bonus inclusions, and free trial vs. money-back guarantee",
            "tone": "Test formal/professional vs. casual/conversational vs. urgent/scarcity-based",
            "audience_segment": "Run the same ad to 2–3 different audience segments to find the most responsive",
            "send_time": "Test different days/times for email or paid post boosts",
        }

        result = {
            "test_design": {
                "content_excerpt": content[:300] + ("…" if len(content) > 300 else ""),
                "variable": variable,
                "num_versions": versions,
                "primary_metric": metric,
                "testing_guidance": guidance.get(variable, "Isolate one clear difference between versions"),
            },
            "version_labels": [
                "Control (A) — your current / baseline version",
                *[f"Variant {chr(66 + i)} — changed version {i + 1}" for i in range(versions - 1)],
            ],
            "methodology": {
                "golden_rule": "Change ONE variable only — everything else stays identical",
                "min_sample_per_version": "500 impressions / 100 opens before drawing conclusions",
                "min_duration": "7 days minimum; 14 days preferred to smooth day-of-week effects",
                "confidence_target": "95% statistical significance before declaring a winner",
                "tool_recommendations": ["Google Optimize (free)", "Klaviyo (email)", "Meta Ads split test tool", "Optimizely"],
            },
            "metrics_to_track": [
                metric,
                "Engagement rate",
                "Conversion rate (not just clicks)",
                "Cost per result (if paid)",
                "Revenue or pipeline value if trackable",
            ],
            "after_the_test": [
                "Implement the winner immediately",
                "Document the result in your 'what works' log",
                "Use the losing variant's insight to form the next hypothesis",
                "Run a new test on the next variable in the funnel",
            ],
        }
        return json.dumps(result, indent=2)

    # ── unknown tool ─────────────────────────────────────────────────────────
    else:
        return json.dumps({"error": f"Unknown tool: {tool_name}"})


# ---------------------------------------------------------------------------
# Agent runner  (streaming generator)
# ---------------------------------------------------------------------------

def run_advertising_agent(
    api_key: str,
    user_request: str,
    business_context: Optional[Dict[str, str]] = None,
) -> Generator[str, None, None]:
    """
    Run the advertising agent in an agentic tool-use loop.

    Yields text chunks (status updates + final markdown response).
    """
    client = anthropic.Anthropic(api_key=api_key)

    user_content = user_request
    if business_context:
        lines = "\n".join(
            f"- **{k}**: {v}" for k, v in business_context.items() if v and str(v).strip()
        )
        if lines:
            user_content = (
                f"**About My Business:**\n{lines}\n\n"
                f"**My Request:**\n{user_request}"
            )

    messages = [{"role": "user", "content": user_content}]
    max_iterations = 12

    for _ in range(max_iterations):
        response = client.messages.create(
            model=ADVERTISING_MODEL,
            max_tokens=8192,
            thinking={"type": "adaptive"},
            system=SYSTEM_PROMPT,
            tools=ADVERTISING_TOOLS,
            messages=messages,
        )

        if response.stop_reason == "end_turn":
            for block in response.content:
                if block.type == "text":
                    yield block.text
            return

        if response.stop_reason == "tool_use":
            tool_blocks = [b for b in response.content if b.type == "tool_use"]
            names = ", ".join(b.name.replace("_", " ") for b in tool_blocks)
            yield f"\n\n---\n⚙️ *Researching: {names}…*\n\n---\n\n"

            messages.append({"role": "assistant", "content": response.content})

            tool_results = [
                {
                    "type": "tool_result",
                    "tool_use_id": tb.id,
                    "content": execute_tool(tb.name, tb.input),
                }
                for tb in tool_blocks
            ]
            messages.append({"role": "user", "content": tool_results})
            continue

        # Unexpected stop reason
        for block in response.content:
            if block.type == "text":
                yield block.text
        return

    yield "\n\n*[Reached maximum iterations — see output above.]*"
